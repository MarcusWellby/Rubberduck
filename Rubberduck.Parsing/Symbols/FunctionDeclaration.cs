﻿using Antlr4.Runtime;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class FunctionDeclaration : Declaration, IDeclarationWithParameter, ICanBeDefaultMember
    {
        private readonly List<ParameterDeclaration> _parameters;

        public FunctionDeclaration(
            QualifiedMemberName name,
            Declaration parent,
            Declaration parentScope,
            string asTypeName,
            VBAParser.AsTypeClauseContext asTypeContext,
            string typeHint,
            Accessibility accessibility,
            ParserRuleContext context,
            Selection selection,
            bool isArray,
            bool isBuiltIn,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(
                  name,
                  parent,
                  parentScope,
                  asTypeName,
                  typeHint,
                  false,
                  false,
                  accessibility,
                  DeclarationType.Function,
                  context,
                  selection,
                  isArray,
                  asTypeContext,
                  isBuiltIn,
                  annotations,
                  attributes)
        {
            _parameters = new List<ParameterDeclaration>();
        }

        public FunctionDeclaration(ComMember member, Declaration parent, QualifiedModuleName module,
            Attributes attributes) : this(
                module.QualifyMemberName(member.Name),
                parent,
                parent,
                member.ReturnType.TypeName,
                null,
                null,
                Accessibility.Global,
                null,
                Selection.Home,
                member.ReturnType.IsArray,
                true,
                null,
                attributes)
        {
            _parameters = member.Parameters.Select(decl => new ParameterDeclaration(decl, this, module)).ToList();
        }

        public IEnumerable<ParameterDeclaration> Parameters
        {
            get
            {
                return _parameters.ToList();
            }
        }

        public void AddParameter(ParameterDeclaration parameter)
        {
            _parameters.Add(parameter);
        }

        /// <summary>
        /// Gets an attribute value indicating whether a member is a class' default member.
        /// If this value is true, any reference to an instance of the class it's the default member of,
        /// should count as a member call to this member.
        /// </summary>
        public bool IsDefaultMember
        {
            get
            {
                IEnumerable<string> value;
                if (Attributes.TryGetValue(IdentifierName + ".VB_UserMemId", out value))
                {
                    return value.Single() == "0";
                }

                return false;
            }
        }
    }
}
