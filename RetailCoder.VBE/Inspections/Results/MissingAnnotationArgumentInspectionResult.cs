using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Inspections.Results
{
    public class MissingAnnotationArgumentInspectionResult : InspectionResultBase
    {
        public MissingAnnotationArgumentInspectionResult(IInspection inspection, string description)
            : base(inspection, description)
        {
        }
    }
}