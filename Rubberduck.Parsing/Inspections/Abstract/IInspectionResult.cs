using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IInspectionResult // todo: rename to IInspectionResult
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
