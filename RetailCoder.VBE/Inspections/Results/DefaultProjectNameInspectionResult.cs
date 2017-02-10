using System;
using System.Collections.Generic;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections.Results
{
    public class DefaultProjectNameInspectionResult : InspectionResultBase
    {
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly RubberduckParserState _state;

        public DefaultProjectNameInspectionResult(IInspection inspection, IInspectionResultTarget target, string name)
            : base(inspection, target, name) { }

        [Obsolete]
        public DefaultProjectNameInspectionResult(IInspection inspection, Declaration target, RubberduckParserState state)
            : base(inspection, target)
        {
            _state = state;
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new RenameProjectQuickFix(Target, Target.QualifiedSelection, Target, _state)
                });
            }
        }

        public override string Description
        {
            get { return Inspection.Description.Captialize(); }
        }
    }
}
