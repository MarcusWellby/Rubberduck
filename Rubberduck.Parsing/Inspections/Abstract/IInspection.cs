using System;
using System.Collections.Generic;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    /// <summary>
    /// An interface that abstracts a runnable code inspection.
    /// </summary>
    public interface IInspection : IInspectionModel, IComparable<IInspection>, IComparable
    {
        /// <summary>
        /// Runs code inspection and returns inspection results.
        /// </summary>
        /// <returns>Returns inspection results, if any.</returns>
        [Obsolete("Use the Execute() method instead.")]
        IEnumerable<IInspectionResult> GetInspectionResults();

        /// <summary>
        /// Runs code inspection and annotates targets.
        /// </summary>
        void Execute();

        /// <summary>
        /// Gets a string that contains additional/meta information about an inspection.
        /// </summary>
        string Meta { get; }
    }
}
