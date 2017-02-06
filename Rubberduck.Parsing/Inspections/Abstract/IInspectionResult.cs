using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IInspectionResult
        : IComparable<IInspectionResult>, IComparable
    {
        string Description { get; }
        string IdentifierName { get; }
        
        IInspection Inspection { get; }
        InspectionResultTarget Target { get; }

        IEnumerable<IQuickFix> QuickFixes { get; }
    }

    public interface ICopyFormatter
    {
        object[] ToArray();
        string ToClipboardString();
    }
}
