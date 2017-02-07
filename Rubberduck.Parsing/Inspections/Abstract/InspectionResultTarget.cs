using System;
using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public enum TargetType
    {
        IDE,
        Project,
        Module,
        Declaration,
        IdentifierReference,
        ParserRuleContext,
    }

    public struct InspectionResultTarget
    {
        private readonly Declaration _module;
        private readonly Selection _selection;

        private readonly object _target;
        private readonly TargetType _type;

        public InspectionResultTarget(ProjectDeclaration target)
            : this(target, Selection.Empty, TargetType.Project, target) { }

        public InspectionResultTarget(ClassModuleDeclaration target)
            : this(target, target.Selection, TargetType.Module, target) { }

        public InspectionResultTarget(ProceduralModuleDeclaration target)
            : this(target, target.Selection, TargetType.Module, target) { }

        public InspectionResultTarget(Declaration module, Declaration target)
            : this(module, target.Selection, TargetType.Declaration, target) { }

        public InspectionResultTarget(Declaration module, IdentifierReference target)
            : this(module, target.Selection, TargetType.IdentifierReference, target) { }

        public InspectionResultTarget(Declaration module, ParserRuleContext target)
            : this(module, target.GetSelection(), TargetType.ParserRuleContext, target) { }

        public InspectionResultTarget(Declaration module, Selection selection, TargetType type, object target = null)
        {
            _module = module;
            _selection = selection;
            _target = target;
            _type = type;
        }

        public Declaration Project
        {
            get
            {
                if (_type == TargetType.Project)
                {
                    return (Declaration) _target;
                }

                return _module != null ? _module.ParentDeclaration : null;
            }
        }

        public Declaration Module { get { return _module; } }

        public ParserRuleContext Context
        {
            get
            {
                switch (_type)
                {
                    case TargetType.IDE:
                        return null;
                    case TargetType.Project:
                    case TargetType.Module:
                    case TargetType.Declaration:
                        return ((Declaration) _target).Context;
                        break;
                    case TargetType.IdentifierReference:
                        return ((IdentifierReference) _target).Context;
                        break;
                    case TargetType.ParserRuleContext:
                        return (ParserRuleContext) _target;
                        break;
                    default:
                        throw new ArgumentOutOfRangeException();
                }
            }
        }

        public Selection Selection { get { return _selection; } }

        public object Target { get { return _target; } }
    }
}