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
        public partial class AnnotationContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class RemCommentContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class IdentifierContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class UntypedIdentifierContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class OptionBaseStmtContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class OptionCompareStmtContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class FileNumberContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class CallStmtContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class ErrorStmtContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class ExpressionContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class DefTypeContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class GoSubStmtContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class GoToStmtContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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

        public partial class WhileWendStmtContext : ICollection<IInspectionResult>
        {
            private readonly InspectionTarget _inspectionTarget = new InspectionTarget();

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
