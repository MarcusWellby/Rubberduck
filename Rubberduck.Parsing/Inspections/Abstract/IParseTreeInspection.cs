using System.Collections.Generic;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IParseTreeInspection : IInspection
    {
        /// <summary>
        /// Parse tree inspections have their results property-injected.
        /// </summary>
        void SetResults(IEnumerable<QualifiedContext> results);
    }
}
