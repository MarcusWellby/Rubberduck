using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Parsing.Grammar
{
    /// <summary>
    /// Provides implementation details for an inspection target.
    /// </summary>
    public class InspectionTarget : ICollection<IInspectionResult>
    {
        private ConcurrentBag<IInspectionResult> _results =
            new ConcurrentBag<IInspectionResult>();

        public IEnumerator<IInspectionResult> GetEnumerator()
        {
            return _results.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        public void Add(IInspectionResult item)
        {
            _results.Add(item);
        }

        public void Clear()
        {
            _results = new ConcurrentBag<IInspectionResult>();
        }

        public bool Contains(IInspectionResult item)
        {
            return _results.Contains(item);
        }

        public void CopyTo(IInspectionResult[] array, int arrayIndex)
        {
            _results.CopyTo(array, arrayIndex);
        }

        public bool Remove(IInspectionResult item)
        {
            throw new NotSupportedException();
        }

        public int Count { get { return _results.Count; } }

        public bool IsReadOnly { get { return false; } }
    }
}