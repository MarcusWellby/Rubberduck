using Antlr4.Runtime;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public struct InspectionResultTarget
    {
        private readonly Declaration _module;
        private readonly Selection _selection;

        private readonly object _target;

        public InspectionResultTarget(Declaration module, Declaration target)
            : this(module, target.Selection)
        {
            _target = target;
        }

        public InspectionResultTarget(Declaration module, IdentifierReference target)
            : this(module, target.Selection)
        {
            _target = target;
        }

        public InspectionResultTarget(Declaration module, ParserRuleContext target)
            : this(module, target.GetSelection())
        {
            _target = target;
        }

        public InspectionResultTarget(Declaration module, Selection selection)
        {
            _module = module;
            _selection = selection;
            _target = null;
        }

        public Declaration Module { get { return _module; } }
        public Selection Selection { get { return _selection; } }
        
        public object Target { get { return _target; } }
    }
}