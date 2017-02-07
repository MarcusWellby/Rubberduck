using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Parsing.Grammar
{
    /// <summary>
    /// An object that can be annotated with inspection results.
    /// </summary>
    public interface IInspectable
    {
        IEnumerable<IInspectionResult> InspectionResults { get; }
        void Annotate(IInspectionResult result);
        void ClearInspectionResults();
    }
}