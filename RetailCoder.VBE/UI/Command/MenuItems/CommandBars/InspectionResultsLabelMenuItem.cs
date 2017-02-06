using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class InspectionResultsLabelMenuItem : CommandMenuItemBase
    {
        public InspectionResultsLabelMenuItem()
            : base(null)
        {
            _caption = string.Empty;
        }

        private string _caption;

        public void SetCaption(IEnumerable<IInspectionResult> inspectionResults)
        {
            _caption = string.Format("{0} issues", inspectionResults.Count());
        }

        public override bool EvaluateCanExecute(RubberduckParserState state)
        {
            return false;
        }

        public override Func<string> Caption { get { return () => _caption; } }
        public override string Key { get { return string.Empty; } }
        public override bool BeginGroup { get { return true; } }
        public override int DisplayOrder { get { return (int)RubberduckCommandBarItemDisplayOrder.InspectionResults; } }
    }
}