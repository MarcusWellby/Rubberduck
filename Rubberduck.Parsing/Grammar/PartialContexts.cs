using System.Collections.Generic;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Parsing.Grammar
{
    public partial class VBAParser
    {
        public partial class AnnotationContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class IdentifierContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class OptionBaseStmtContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class OptionCompareStmtContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class FileNumberContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class CallStmtContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class ErrorStmtContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class ExpressionContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class DefTypeContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class GoSubStmtContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class GoToStmtContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class WhileWendStmtContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }

        public partial class RemCommentContext : IInspectable
        {
            private readonly InspectableContext _inspectable = new InspectableContext();
            public IEnumerable<IInspectionResult> InspectionResults { get { return _inspectable.InspectionResults; } }
            public void Annotate(IInspectionResult result) { _inspectable.Annotate(result); }
            public void ClearInspectionResults() { _inspectable.ClearInspectionResults(); }
        }
    }
}
