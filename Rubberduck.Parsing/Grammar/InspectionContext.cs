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
            private readonly ConcurrentBag<IInspectionResult> _results = 
                new ConcurrentBag<IInspectionResult>();

            public IEnumerable<IInspectionResult> InspectionResults { get { return _results; } }

            public void Annotate(IInspectionResult result)
            {
                _results.Add(result);
            }

            public void ClearInspectionResults()
            {
                while (!_results.IsEmpty)
                {
                    IInspectionResult item;
                    _results.TryTake(out item);
                }
            }
        }
    }
}
