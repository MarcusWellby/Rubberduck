using NLog;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.Binding;
using Rubberduck.VBEditor;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.VBEditor.Application;

namespace Rubberduck.Parsing.Symbols
{
    internal static class DictionaryExtensions
    {
        public static IEnumerable<TValue> AllValues<TKey, TValue>(
            this ConcurrentDictionary<TKey, ConcurrentBag<TValue>> source)
        {
            return source.SelectMany(item => item.Value).ToList();
        }

        public static IEnumerable<TValue> AllValues<TKey, TValue>(
        this IDictionary<TKey, IList<TValue>> source)
        {
            return source.SelectMany(item => item.Value).ToList();
        }

        public static ConcurrentDictionary<TKey, ConcurrentBag<TValue>> ToConcurrentDictionary<TKey, TValue>(this IEnumerable<IGrouping<TKey, TValue>> source)
        {
            return new ConcurrentDictionary<TKey, ConcurrentBag<TValue>>(source.Select(x => new KeyValuePair<TKey, ConcurrentBag<TValue>>(x.Key, new ConcurrentBag<TValue>(x))));
        }   
    }

    public class DeclarationFinder
    {
        private static readonly SquareBracketedNameComparer NameComparer = new SquareBracketedNameComparer();
        private static readonly Logger Logger = LogManager.GetCurrentClassLogger();

        private readonly IHostApplication _hostApp;
        private readonly AnnotationService _annotationService;
        private readonly ConcurrentDictionary<string, ConcurrentBag<Declaration>> _declarationsByName;
        private readonly ConcurrentDictionary<QualifiedModuleName, ConcurrentBag<Declaration>> _declarations;
        private readonly ConcurrentDictionary<QualifiedMemberName, ConcurrentBag<Declaration>> _undeclared;
        private readonly ConcurrentBag<UnboundMemberDeclaration> _unresolved;
        private readonly ConcurrentDictionary<QualifiedModuleName, ConcurrentBag<IAnnotation>> _annotations;
        private readonly ConcurrentDictionary<Declaration, ConcurrentBag<Declaration>> _parametersByParent;
        private readonly ConcurrentDictionary<DeclarationType, ConcurrentBag<Declaration>> _userDeclarationsByType;

        private readonly Lazy<ConcurrentDictionary<Declaration, Declaration[]>> _handlersByWithEventsField;
        private readonly Lazy<ConcurrentDictionary<VBAParser.ImplementsStmtContext, Declaration[]>> _membersByImplementsContext;
        private readonly Lazy<ConcurrentDictionary<Declaration, Declaration[]>> _interfaceMembers;
        
        private static readonly object ThreadLock = new object();

        public DeclarationFinder(IReadOnlyList<Declaration> declarations, IEnumerable<IAnnotation> annotations, IHostApplication hostApp = null)
        {
            _hostApp = hostApp;
            _annotations = annotations.GroupBy(node => node.QualifiedSelection.QualifiedName).ToConcurrentDictionary();
            _declarations = declarations.GroupBy(item => item.QualifiedName.QualifiedModuleName).ToConcurrentDictionary();
            _declarationsByName = declarations.GroupBy(declaration => declaration.IdentifierName.ToLowerInvariant()).ToConcurrentDictionary();
            _parametersByParent = declarations.Where(declaration => declaration.DeclarationType == DeclarationType.Parameter)
                .GroupBy(declaration => declaration.ParentDeclaration).ToConcurrentDictionary();
            _userDeclarationsByType = declarations.Where(declaration => !declaration.IsBuiltIn).GroupBy(declaration => declaration.DeclarationType).ToConcurrentDictionary();
            _builtinEvents = new Lazy<ConcurrentBag<Declaration>>(() => FindBuiltInEventHandlers(declarations));

            _projects = _projects = new Lazy<ConcurrentBag<Declaration>>(() => new ConcurrentBag<Declaration>(declarations.Where(d => d.DeclarationType == DeclarationType.Project)));
            _classes = new Lazy<ConcurrentBag<Declaration>>(() => new ConcurrentBag<Declaration>(declarations.Where(d => d.DeclarationType == DeclarationType.ClassModule)));

            var withEventsFields = UserDeclarations(DeclarationType.Variable).Where(item => item.IsWithEvents).ToArray();
            var events = withEventsFields.Select(field =>
                new
                {
                    WithEventsField = field,
                    AvailableEvents = FindEvents(field.AsTypeDeclaration).ToArray()
                });

            _handlersByWithEventsField = new Lazy<ConcurrentDictionary<Declaration, Declaration[]>>(() =>
                new ConcurrentDictionary<Declaration, Declaration[]>(
                    events.Select(item =>
                        new
                        {
                            item.WithEventsField,
                            Handlers = item.AvailableEvents.SelectMany(evnt =>
                                _declarations[item.WithEventsField.ParentDeclaration.QualifiedName.QualifiedModuleName]
                                    .Where(member => member.DeclarationType == DeclarationType.Procedure
                                                  && member.IdentifierName == item.WithEventsField.IdentifierName + "_" + evnt.IdentifierName))
                        })
                        .ToDictionary(item => item.WithEventsField, item => item.Handlers.ToArray())
                    ));

            _undeclared = new ConcurrentDictionary<QualifiedMemberName, ConcurrentBag<Declaration>>(new Dictionary<QualifiedMemberName, ConcurrentBag<Declaration>>());
            _unresolved = new ConcurrentBag<UnboundMemberDeclaration>(new List<UnboundMemberDeclaration>());
            _annotationService = new AnnotationService(this);

            var implementsInstructions = UserDeclarations(DeclarationType.ClassModule).SelectMany(cls => 
                cls.References.Where(reference => ParserRuleContextHelper.HasParent<VBAParser.ImplementsStmtContext>(reference.Context))
                    .Select(reference => new { IdentifierReference = reference, Context = ParserRuleContextHelper.GetParent<VBAParser.ImplementsStmtContext>(reference.Context)}));

            var interfaceMembers = implementsInstructions.Select(item => new
                {
                    InterfaceModule = item.IdentifierReference.Declaration,
                    InterfaceMembers = _declarations[item.IdentifierReference.Declaration.QualifiedName.QualifiedModuleName]
                    .Where(member => member.DeclarationType.HasFlag(DeclarationType.Member))
                });

            _interfaceMembers = new Lazy<ConcurrentDictionary<Declaration, Declaration[]>>(() =>
               new ConcurrentDictionary<Declaration, Declaration[]>(interfaceMembers.ToDictionary(item => item.InterfaceModule, item => item.InterfaceMembers.ToArray())));

            var implementingNames = new Lazy<IEnumerable<string>>(() => implementsInstructions.SelectMany(item =>
                    _declarations[item.IdentifierReference.Declaration.QualifiedName.QualifiedModuleName]
                        .Where(member => member.DeclarationType.HasFlag(DeclarationType.Member))
                        .Select(member => item.IdentifierReference.Declaration.IdentifierName + "_" + member.IdentifierName)));

            var implementableMembers = implementsInstructions.Select(item =>
                new
                {
                    item.Context,
                    Members = _declarations[item.IdentifierReference.QualifiedModuleName].Where(implementingTypeMember =>
                        implementingNames.Value.Contains(implementingTypeMember.IdentifierName)).ToArray()
                });

            _membersByImplementsContext = new Lazy<ConcurrentDictionary<VBAParser.ImplementsStmtContext, Declaration[]>>(() =>
            new ConcurrentDictionary<VBAParser.ImplementsStmtContext, Declaration[]>(
                implementableMembers.ToDictionary(item => item.Context, item => item.Members)));
        }

        public IEnumerable<Declaration> Undeclared
        {
            get { return _undeclared.AllValues(); }
        }

        private IEnumerable<Declaration> _nonBaseAsType;
        public IEnumerable<Declaration> FindDeclarationsWithNonBaseAsType()
        {
            lock (ThreadLock)
            {
                return _nonBaseAsType ?? (_nonBaseAsType = _declarations.AllValues().Where(d =>
                            !string.IsNullOrWhiteSpace(d.AsTypeName)
                            && !d.AsTypeIsBaseType
                            && d.DeclarationType != DeclarationType.Project
                            && d.DeclarationType != DeclarationType.ProceduralModule).ToList());
            }
        }

        private readonly Lazy<ConcurrentBag<Declaration>> _builtinEvents; 
        public IEnumerable<Declaration> FindBuiltinEventHandlers()
        {
            lock (ThreadLock)
            {
                return _builtinEvents.Value;
            }
        }

        private readonly Lazy<ConcurrentBag<Declaration>> _classes;

        public IEnumerable<Declaration> Classes
        {
            get
            {
                lock (ThreadLock)
                {
                    return _classes.Value;
                }
            }
        }

        private readonly Lazy<ConcurrentBag<Declaration>> _projects;

        public IEnumerable<Declaration> Projects
        {
            get
            {
                lock (ThreadLock)
                {
                    return _projects.Value;
                }
            }
        }

        public IEnumerable<Declaration> UserDeclarations(DeclarationType type)
        {
            ConcurrentBag<Declaration> result;
            if (!_userDeclarationsByType.TryGetValue(type, out result))
            {
                result = new ConcurrentBag<Declaration>(_userDeclarationsByType
                    .Where(item => item.Key.HasFlag(type))
                    .SelectMany(item => item.Value));
            }
            return result;
        }

        public IEnumerable<UnboundMemberDeclaration> UnresolvedMemberDeclarations()
        {
            lock (ThreadLock)
            {
                return _unresolved.ToArray();
            }            
        }

        public IEnumerable<Declaration> FindHandlersForWithEventsField(Declaration field)
        {
            Declaration[] result;
            return _handlersByWithEventsField.Value.TryGetValue(field, out result) 
                ? result 
                : Enumerable.Empty<Declaration>();
        }

        public IEnumerable<Declaration> FindInterfaceMembersForImplementsContext(VBAParser.ImplementsStmtContext context)
        {
            Declaration[] result;
            return _membersByImplementsContext.Value.TryGetValue(context, out result)
                ? result
                : Enumerable.Empty<Declaration>();
        }

        public IEnumerable<Declaration> FindAllInterfaceMembers()
        {
            return _interfaceMembers.Value.SelectMany(item => item.Value);
        }

        public IEnumerable<Declaration> FindAllInterfaceImplementingMembers()
        {
            return _membersByImplementsContext.Value.SelectMany(item => item.Value);
        }

        public Declaration FindParameter(Declaration procedure, string parameterName)
        {
            return _parametersByParent[procedure].SingleOrDefault(parameter => parameter.IdentifierName == parameterName);
        }

        public IEnumerable<Declaration> FindMemberMatches(Declaration parent, string memberName)
        {
            ConcurrentBag<Declaration> children;
            if (_declarations.TryGetValue(parent.QualifiedName.QualifiedModuleName, out children))
            {
                return children.Where(item => item.DeclarationType.HasFlag(DeclarationType.Member)
                                             && item.IdentifierName == memberName).ToList();
            }

            return Enumerable.Empty<Declaration>();
        }

        public IEnumerable<IAnnotation> FindAnnotations(QualifiedModuleName module)
        {
            ConcurrentBag<IAnnotation> result;
            return _annotations.TryGetValue(module, out result) ? result : Enumerable.Empty<IAnnotation>();
        }

        public bool IsMatch(string declarationName, string potentialMatchName)
        {
            return string.Equals(declarationName, potentialMatchName, StringComparison.OrdinalIgnoreCase);
        }

        private IEnumerable<Declaration> FindEvents(Declaration module)
        {
            Debug.Assert(module != null);

            var members = _declarations[module.QualifiedName.QualifiedModuleName];
            return members == null 
                ? Enumerable.Empty<Declaration>() 
                : members.Where(declaration => declaration.DeclarationType == DeclarationType.Event).ToList();
        }

        public Declaration FindEvent(Declaration module, string eventName)
        {
            var matches = MatchName(eventName);
            return matches.FirstOrDefault(m => module.Equals(Declaration.GetModuleParent(m)) && m.DeclarationType == DeclarationType.Event);
        }

        public Declaration FindLabel(Declaration procedure, string label)
        {
            var matches = MatchName(label);
            return matches.FirstOrDefault(m => procedure.Equals(m.ParentDeclaration) && m.DeclarationType == DeclarationType.LineLabel);
        }

        public IEnumerable<Declaration> MatchName(string name)
        {
            var normalizedName = ToNormalizedName(name);
            ConcurrentBag<Declaration> result;
            return _declarationsByName.TryGetValue(normalizedName, out result) 
                ? result 
                : Enumerable.Empty<Declaration>();
        }

        private string ToNormalizedName(string name)
        {
            var lower = name.ToLowerInvariant();
            if (lower.Length > 1 && lower[0] == '[' && lower[lower.Length - 1] == ']')
            {
                var result = lower.Substring(1, lower.Length - 2);
                return result;
            }
            return lower;
        }

        public Declaration FindProject(string name, Declaration currentScope = null)
        {
            Declaration result = null;
            try
            {
                result = MatchName(name).SingleOrDefault(project => project.DeclarationType.HasFlag(DeclarationType.Project)
                    && (currentScope == null || project.ProjectId == currentScope.ProjectId));
            }
            catch (InvalidOperationException exception)
            {
                Logger.Error(exception, "Multiple matches found for project '{0}'.", name);
            }

            return result;
        }

        public Declaration FindStdModule(string name, Declaration parent = null, bool includeBuiltIn = false)
        {
            Debug.Assert(parent != null);
            Declaration result = null;
            try
            {
                var matches = MatchName(name);
                result = matches.SingleOrDefault(declaration => declaration.DeclarationType.HasFlag(DeclarationType.ProceduralModule)
                    && (parent.Equals(declaration.ParentDeclaration))
                    && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (InvalidOperationException exception)
            {
                Logger.Error(exception, "Multiple matches found for std.module '{0}'.", name);
            }

            return result;
        }

        public Declaration FindClassModule(string name, Declaration parent = null, bool includeBuiltIn = false)
        {
            Debug.Assert(parent != null);
            Declaration result = null;
            try
            {
                var matches = MatchName(name);
                result = matches.SingleOrDefault(declaration => declaration.DeclarationType.HasFlag(DeclarationType.ClassModule)
                    && (parent.Equals(declaration.ParentDeclaration))
                    && (includeBuiltIn || !declaration.IsBuiltIn));
            }
            catch (InvalidOperationException exception)
            {
                Logger.Error(exception, "Multiple matches found for class module '{0}'.", name);
            }

            return result;
        }

        public Declaration FindReferencedProject(Declaration callingProject, string referencedProjectName)
        {
            return FindInReferencedProjectByPriority(callingProject, referencedProjectName, p => p.DeclarationType.HasFlag(DeclarationType.Project));
        }

        public Declaration FindModuleEnclosingProjectWithoutEnclosingModule(Declaration callingProject, Declaration callingModule, string calleeModuleName, DeclarationType moduleType)
        {
            var nameMatches = MatchName(calleeModuleName);
            var moduleMatches = nameMatches.Where(m =>
                m.DeclarationType.HasFlag(moduleType)
                && Declaration.GetProjectParent(m).Equals(callingProject)
                && !m.Equals(callingModule));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassEnclosingProject(Declaration callingProject, Declaration callingModule, string defaultInstanceVariableClassName)
        {
            var nameMatches = MatchName(defaultInstanceVariableClassName);
            var moduleMatches = nameMatches.Where(m =>
                m.DeclarationType == DeclarationType.ClassModule && ((ClassModuleDeclaration)m).HasDefaultInstanceVariable
                && Declaration.GetProjectParent(m).Equals(callingProject)).ToList(); 
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindModuleReferencedProject(Declaration callingProject, Declaration callingModule, string calleeModuleName, DeclarationType moduleType)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName, p => p.DeclarationType.HasFlag(moduleType));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindModuleReferencedProject(Declaration callingProject, Declaration callingModule, Declaration referencedProject, string calleeModuleName, DeclarationType moduleType)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName, p => referencedProject.Equals(Declaration.GetProjectParent(p)) && p.DeclarationType.HasFlag(moduleType));
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassReferencedProject(Declaration callingProject, Declaration callingModule, string calleeModuleName)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName, p => p.DeclarationType == DeclarationType.ClassModule && ((ClassModuleDeclaration)p).HasDefaultInstanceVariable);
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindDefaultInstanceVariableClassReferencedProject(Declaration callingProject, Declaration callingModule, Declaration referencedProject, string calleeModuleName)
        {
            var moduleMatches = FindAllInReferencedProjectByPriority(callingProject, calleeModuleName,
                p => referencedProject.Equals(Declaration.GetProjectParent(p))
                    && p.DeclarationType == DeclarationType.ClassModule 
                    && ((ClassModuleDeclaration)p).HasDefaultInstanceVariable);
            var accessibleModules = moduleMatches.Where(calledModule => AccessibilityCheck.IsModuleAccessible(callingProject, callingModule, calledModule));
            var match = accessibleModules.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberWithParent(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration parent, string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches
                .Where(m => m.DeclarationType.HasFlag(memberType)
                            && parent.Equals(m.ParentDeclaration))
                .ToList();
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            if (match != null)
            {
                return match;
            }
            return ClassModuleDeclaration
                .GetSupertypes(parent)
                .Select(supertype => 
                    FindMemberWithParent(callingProject, callingModule, callingParent, supertype, memberName, memberType))
                .FirstOrDefault(supertypeMember => supertypeMember != null);
        }

        public Declaration FindMemberEnclosingModule(Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            // We do not explicitly pass the callingProject here because we have to walk up the type hierarchy
            // and thus the project differs depending on the callingModule.
            var callingProject = Declaration.GetProjectParent(callingModule);
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches
                .Where(m => m.DeclarationType.HasFlag(memberType)
                            && Declaration.GetProjectParent(m).Equals(callingProject)
                            && callingModule.Equals(Declaration.GetModuleParent(m))
                ).ToList();
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            if (match != null)
            {
                return match;
            }
            // Classes such as Worksheet have properties such as Range that can be access in a user defined class such as Sheet1,
            // that's why we have to walk the type hierarchy and find these implementations.
            foreach (var supertype in ClassModuleDeclaration.GetSupertypes(callingModule))
            {
                // Only built-in classes such as Worksheet can be considered "real base classes".
                // User created interfaces work differently and don't allow accessing accessing implementations.
                if (!supertype.IsBuiltIn)
                {
                    continue;
                }
                var supertypeMatch = FindMemberEnclosingModule(supertype, callingParent, memberName, memberType);
                if (supertypeMatch != null)
                {
                    return supertypeMatch;
                }
            }

            return null;
        }

        public Declaration FindMemberEnclosingProcedure(Declaration enclosingProcedure, string memberName, DeclarationType memberType, ParserRuleContext onSiteContext = null)
        {
            if (memberType == DeclarationType.Variable && NameComparer.Equals(enclosingProcedure.IdentifierName, memberName))
            {
                return enclosingProcedure;
            }
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m =>
                m.DeclarationType.HasFlag(memberType)
                && enclosingProcedure.Equals(m.ParentDeclaration));
            return memberMatches.FirstOrDefault();
        }

        public Declaration OnUndeclaredVariable(Declaration enclosingProcedure, string identifierName, ParserRuleContext context)
        {
            var annotations = _annotationService.FindAnnotations(enclosingProcedure.QualifiedName.QualifiedModuleName, context.Start.Line);
            var undeclaredLocal =
                new Declaration(
                    new QualifiedMemberName(enclosingProcedure.QualifiedName.QualifiedModuleName, identifierName),
                    enclosingProcedure, enclosingProcedure, "Variant", string.Empty, false, false,
                    Accessibility.Implicit, DeclarationType.Variable, context, context.GetSelection(), false, null,
                    false, annotations, null, true);

            var hasUndeclared = _undeclared.ContainsKey(enclosingProcedure.QualifiedName);
            if (hasUndeclared)
            {
                ConcurrentBag<Declaration> undeclared;
                while (!_undeclared.TryGetValue(enclosingProcedure.QualifiedName, out undeclared))
                {
                    _undeclared.TryGetValue(enclosingProcedure.QualifiedName, out undeclared);
                }
                var inScopeUndeclared = undeclared.FirstOrDefault(d => d.IdentifierName == identifierName);
                if (inScopeUndeclared != null)
                {
                    return inScopeUndeclared;
                }
                undeclared.Add(undeclaredLocal);
            }
            else
            {
                _undeclared.TryAdd(enclosingProcedure.QualifiedName, new ConcurrentBag<Declaration> { undeclaredLocal });
            }
            return undeclaredLocal;
        }

        public void AddUnboundContext(Declaration parentDeclaration, VBAParser.LExprContext context, IBoundExpression withExpression)
        {
            
            //The only forms we care about right now are MemberAccessExprContext or WithMemberAccessExprContext.
            var access = (ParserRuleContext)context.GetChild<VBAParser.MemberAccessExprContext>(0)
                         ?? context.GetChild<VBAParser.WithMemberAccessExprContext>(0);
            if (access == null)
            {
                return;
            }

            var identifier = access.GetChild<VBAParser.UnrestrictedIdentifierContext>(0);
            var annotations = _annotationService.FindAnnotations(parentDeclaration.QualifiedName.QualifiedModuleName, context.Start.Line);

            var declaration = new UnboundMemberDeclaration(parentDeclaration, identifier,
                (access is VBAParser.MemberAccessExprContext) ? (ParserRuleContext)access.children[0] : withExpression.Context, 
                annotations);

            _unresolved.Add(declaration);
        }

        public Declaration OnBracketedExpression(string expression, ParserRuleContext context)
        {
            var hostApp = FindProject(_hostApp == null ? "VBA" : _hostApp.ApplicationName);
            Debug.Assert(hostApp != null, "Host application project can't be null. Make sure VBA standard library is included if host is unknown.");

            var qualifiedName = hostApp.QualifiedName.QualifiedModuleName.QualifyMemberName(expression);

            ConcurrentBag<Declaration> undeclared;
            if (_undeclared.TryGetValue(qualifiedName, out undeclared))
            {
                return undeclared.SingleOrDefault();
            }

            var item = new Declaration(qualifiedName, hostApp, hostApp, Tokens.Variant, string.Empty, false, false, Accessibility.Global, DeclarationType.BracketedExpression, context, context.GetSelection(), false, null);
            _undeclared.TryAdd(qualifiedName, new ConcurrentBag<Declaration> { item });
            return item;
        }

        public Declaration FindMemberEnclosedProjectWithoutEnclosingModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches.Where(m => m.DeclarationType.HasFlag(memberType)
                && (Declaration.GetModuleParent(m).DeclarationType == DeclarationType.ProceduralModule 
                    || memberType == DeclarationType.Enumeration 
                    || memberType == DeclarationType.EnumerationMember)
                && Declaration.GetProjectParent(m).Equals(callingProject)
                && !callingModule.Equals(Declaration.GetModuleParent(m)))
                .ToList();
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        private static bool IsInstanceSensitive(DeclarationType memberType)
        {
            return memberType.HasFlag(DeclarationType.Procedure)
                || memberType.HasFlag(DeclarationType.Function) 
                || memberType.HasFlag(DeclarationType.Variable)
                || memberType.HasFlag(DeclarationType.Constant);
        }

        public Declaration FindMemberEnclosedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration memberModule, string memberName, DeclarationType memberType)
        {
            var allMatches = MatchName(memberName);
            var memberMatches = allMatches
                .Where(m => m.DeclarationType.HasFlag(memberType)
                            && Declaration.GetProjectParent(m).Equals(callingProject)
                            && memberModule.Equals(Declaration.GetModuleParent(m)))
                .ToList();

            var match = memberMatches.FirstOrDefault(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            if (match != null)
            {
                return match;
            }

            return ClassModuleDeclaration
                .GetSupertypes(memberModule)
                .Select(supertype => 
                    FindMemberEnclosedProjectInModule(callingProject, callingModule, callingParent, supertype, memberName, memberType))
                .FirstOrDefault(supertypeMember => supertypeMember != null);
        }

        public Declaration FindMemberReferencedProject(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            var isInstanceSensitive = IsInstanceSensitive(memberType);
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => (!isInstanceSensitive || Declaration.GetModuleParent(p) == null || Declaration.GetModuleParent(p).DeclarationType != DeclarationType.ClassModule) && p.DeclarationType.HasFlag(memberType));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, DeclarationType moduleType, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && (Declaration.GetModuleParent(p) == null || Declaration.GetModuleParent(p).DeclarationType == moduleType));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInGlobalClassModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && (Declaration.GetModuleParent(p) == null || Declaration.GetModuleParent(p).DeclarationType == DeclarationType.ClassModule) && ((ClassModuleDeclaration)Declaration.GetModuleParent(p)).IsGlobalClassModule);
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            return match;
        }

        public Declaration FindMemberReferencedProjectInModule(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration memberModule, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && memberModule.Equals(Declaration.GetModuleParent(p)));
            var accessibleMembers = memberMatches.Where(m => AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
            var match = accessibleMembers.FirstOrDefault();
            if (match != null)
            {
                return match;
            }
            return ClassModuleDeclaration
                .GetSupertypes(memberModule)
                .Select(supertype => 
                    FindMemberReferencedProjectInModule(callingProject, callingModule, callingParent, supertype, memberName, memberType))
                .FirstOrDefault(supertypeMember => supertypeMember != null);
        }

        public Declaration FindMemberReferencedProject(Declaration callingProject, Declaration callingModule, Declaration callingParent, Declaration referencedProject, string memberName, DeclarationType memberType)
        {
            var memberMatches = FindAllInReferencedProjectByPriority(callingProject, memberName, p => p.DeclarationType.HasFlag(memberType) && referencedProject.Equals(Declaration.GetProjectParent(p)));
            return memberMatches.FirstOrDefault(m => 
                    AccessibilityCheck.IsMemberAccessible(callingProject, callingModule, callingParent, m));
        }

        private Declaration FindInReferencedProjectByPriority(Declaration enclosingProject, string name, Func<Declaration, bool> predicate)
        {
            return FindAllInReferencedProjectByPriority(enclosingProject, name, predicate).FirstOrDefault();
        }

        private IEnumerable<Declaration> FindAllInReferencedProjectByPriority(Declaration enclosingProject, string name, Func<Declaration, bool> predicate)
        {
            var interprojectMatches = MatchName(name).Where(predicate).ToList();
            var projectReferences = ((ProjectDeclaration)enclosingProject).ProjectReferences.ToList();
            if (interprojectMatches.Count == 0)
            {
                yield break;
            }
            foreach (var projectReference in projectReferences)
            {
                var match = interprojectMatches.FirstOrDefault(interprojectMatch => interprojectMatch.ProjectId == projectReference.ReferencedProjectId);
                if (match != null)
                {
                    yield return match;
                }
            }
        }

        public ConcurrentBag<Declaration> FindBuiltInEventHandlers(IEnumerable<Declaration> declarations)
        {
            var declarationList = declarations.ToList();

            var handlerNames = declarationList.Where(declaration => declaration.IsBuiltIn && declaration.DeclarationType == DeclarationType.Event)
                                           .SelectMany(e =>
                                           {
                                               var parentModuleSubtypes = ((ClassModuleDeclaration)e.ParentDeclaration).Subtypes;
                                               return parentModuleSubtypes.Any()
                                                   ? parentModuleSubtypes.Select(v => v.IdentifierName + "_" + e.IdentifierName)
                                                   : new[] { e.ParentDeclaration.IdentifierName + "_" + e.IdentifierName };
                                           });

            var handlers = declarationList.Where(item =>
                // class module built-in events
                (item.DeclarationType == DeclarationType.Procedure &&
                 item.ParentDeclaration.DeclarationType == DeclarationType.ClassModule && (
                     item.IdentifierName.Equals("Class_Initialize", StringComparison.InvariantCultureIgnoreCase) ||
                     item.IdentifierName.Equals("Class_Terminate", StringComparison.InvariantCultureIgnoreCase))) ||
                // standard module built-in handlers (Excel specific):
                (_hostApp != null &&
                 _hostApp.ApplicationName.Equals("Excel", StringComparison.InvariantCultureIgnoreCase) &&
                 item.DeclarationType == DeclarationType.Procedure &&
                 item.ParentDeclaration.DeclarationType == DeclarationType.ProceduralModule && (
                     item.IdentifierName.Equals("auto_open", StringComparison.InvariantCultureIgnoreCase) ||
                     item.IdentifierName.Equals("auto_close", StringComparison.InvariantCultureIgnoreCase))) ||
                // user handlers:
                (!item.IsBuiltIn &&
                 item.DeclarationType == DeclarationType.Procedure &&
                 handlerNames.Contains(item.IdentifierName))
                );

            return new ConcurrentBag<Declaration>(handlers);
        }
    }
}
