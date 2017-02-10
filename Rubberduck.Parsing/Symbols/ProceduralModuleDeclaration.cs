using System;
using System.Collections;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProceduralModuleDeclaration : Declaration, IInspectionResultTarget<ProceduralModuleDeclaration>
    {
        public ProceduralModuleDeclaration(
            QualifiedMemberName qualifiedName,
            Declaration projectDeclaration,
            string name,
            bool isBuiltIn,
            IEnumerable<IAnnotation> annotations,
            Attributes attributes)
            : base(
                  qualifiedName,
                  projectDeclaration,
                  projectDeclaration,
                  name,
                  null,
                  false,
                  false,
                  Accessibility.Public,
                  DeclarationType.ProceduralModule,
                  null,
                  Selection.Home,
                  false,
                  null,
                  isBuiltIn,
                  annotations,
                  attributes) { }

        public ProceduralModuleDeclaration(ComModule statics, Declaration parent, QualifiedModuleName module,
            Attributes attributes)
            : this(
                module.QualifyMemberName(statics.Name),
                parent,
                statics.Name,
                true,
                new List<IAnnotation>(),
                attributes)
        {
            IsPrivateModule = statics.IsRestricted;
        }

        //These are the pseudo-module ctor for COM enumerations and types.
        public ProceduralModuleDeclaration(ComEnumeration pseudo, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName(string.Format("_{0}", pseudo.Name)),
                parent,
                string.Format("_{0}", pseudo.Name),
                true,
                new List<IAnnotation>(),
                new Attributes()) { }

        public ProceduralModuleDeclaration(ComStruct pseudo, Declaration parent, QualifiedModuleName module)
            : this(
                module.QualifyMemberName(string.Format("_{0}", pseudo.Name)),
                parent,
                string.Format("_{0}", pseudo.Name),
                true,
                new List<IAnnotation>(),
                new Attributes()) { }

        public bool IsPrivateModule { get; internal set; }

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

        public ProceduralModuleDeclaration Target { get { return this; } }
    }
}
