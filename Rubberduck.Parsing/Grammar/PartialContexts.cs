using System;
using System.Collections;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Parsing.Grammar
{
    /// <summary>
    /// Provides annotation layer for the <see cref="ParserRuleContext"/> types nested under <see cref="VBAParser"/>.
    /// </summary>
    public partial class VBAParser
    {
        public partial class AnnotationContext : IInspectionResultTarget<AnnotationContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public AnnotationContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class RemCommentContext : IInspectionResultTarget<RemCommentContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public RemCommentContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class IdentifierContext : IInspectionResultTarget<IdentifierContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public IdentifierContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class UntypedIdentifierContext : IInspectionResultTarget<UntypedIdentifierContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public UntypedIdentifierContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class OptionBaseStmtContext : IInspectionResultTarget<OptionBaseStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public OptionBaseStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class OptionCompareStmtContext : IInspectionResultTarget<OptionCompareStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public OptionCompareStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class FileNumberContext : IInspectionResultTarget<FileNumberContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public FileNumberContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class CallStmtContext : IInspectionResultTarget<CallStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public CallStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class ErrorStmtContext : IInspectionResultTarget<ErrorStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public ErrorStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class ExpressionContext : IInspectionResultTarget<ExpressionContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public ExpressionContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class DefTypeContext : IInspectionResultTarget<DefTypeContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public DefTypeContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class GoSubStmtContext : IInspectionResultTarget<GoSubStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public GoSubStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class GoToStmtContext : IInspectionResultTarget<GoToStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public GoToStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class WhileWendStmtContext : IInspectionResultTarget<WhileWendStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public WhileWendStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class ForNextStmtContext : IInspectionResultTarget<ForNextStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public ForNextStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }

        public partial class IfStmtContext : IInspectionResultTarget<IfStmtContext>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();
            public IfStmtContext Target { get { return this; } }

            #region ICollection<IInspectionResult>
            public IEnumerator<IInspectionResult> GetEnumerator()
            {
                return _inspectionTarget.GetEnumerator();
            }

            IEnumerator IEnumerable.GetEnumerator()
            {
                return GetEnumerator();
            }

            public void Add(IInspectionResult item)
            {
                _inspectionTarget.Add(item);
            }

            public void Clear()
            {
                _inspectionTarget.Clear();
            }

            public bool Contains(IInspectionResult item)
            {
                return _inspectionTarget.Contains(item);
            }

            public void CopyTo(IInspectionResult[] array, int arrayIndex)
            {
                _inspectionTarget.CopyTo(array, arrayIndex);
            }

            [Obsolete("Throws NotSupportedException. Use Clear() method.")]
            public bool Remove(IInspectionResult item)
            {
                return false;
            }

            public int Count { get { return _inspectionTarget.Count; } }

            public bool IsReadOnly { get { return _inspectionTarget.IsReadOnly; } }
            #endregion
        }
    }
}
