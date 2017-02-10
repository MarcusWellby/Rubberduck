using System;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    /// <summary>
    /// Describes an inspection result.
    /// </summary>
    public interface IInspectionResult //: IComparable<IInspectionResult>, IComparable
    {
        /// <summary>
        /// The target (parent) object.
        /// </summary>
        IInspectionResultTarget Target { get; }

        /// <summary>
        /// The localized result description.
        /// </summary>
        string Description { get; }
        
        /// <summary>
        /// The inspection that produced this result.
        /// </summary>
        IInspection Inspection { get; }

        /// <summary>
        /// The <see cref="IQuickFix"/> options available.
        /// </summary>
        IEnumerable<IQuickFix> QuickFixes { get; }

        /// <summary>
        /// The default <see cref="IQuickFix"/> for this result, if any.
        /// </summary>
        IQuickFix DefaultQuickFix { get; }
    }

 

    public interface ICopyFormatter
    {
        object[] ToArray();
        string ToClipboardString();
    }
}
