using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Parsing.Grammar
{
    public interface IInspectable
    {
        IEnumerable<IInspectionResult> InspectionResults { get; }
        void Annotate(IInspectionResult result);
        void ClearInspectionResults();
    }
}