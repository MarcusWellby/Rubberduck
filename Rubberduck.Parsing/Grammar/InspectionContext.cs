using System;
using System.Collections.Concurrent;
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

    public partial class VBAParser
    {
        public partial class AnnotationContext : IInspectable
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
}
