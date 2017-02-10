using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Parsing.Grammar
{
    public interface IInspectionResultTarget : ICollection<IInspectionResult>
    {
    }

    public interface IInspectionResultTarget<out TTarget> : IInspectionResultTarget
    {
        TTarget Target { get; }
    }
}