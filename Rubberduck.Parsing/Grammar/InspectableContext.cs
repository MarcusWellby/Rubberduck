using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Parsing.Grammar
{
    /// <summary>
    /// Provides implementation details for <see cref="IInspectable"/> interface.
    /// </summary>
    public class InspectableContext : IInspectable
    {
        private ConcurrentBag<IInspectionResult> _results =
            new ConcurrentBag<IInspectionResult>();

        public IEnumerable<IInspectionResult> InspectionResults { get { return _results; } }

        public void Annotate(IInspectionResult result)
        {
            try
            {
                _results.Add(result);
            }
            catch (Exception e)
            {
                System.Diagnostics.Debug.Assert(false, e.Message, e.ToString());
            }
        }

        public void ClearInspectionResults()
        {
            _results = new ConcurrentBag<IInspectionResult>();
        }
    }
}