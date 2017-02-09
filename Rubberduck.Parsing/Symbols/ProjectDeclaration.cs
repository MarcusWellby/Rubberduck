using System;
using System.Collections;
using Rubberduck.Parsing.ComReflection;
using Rubberduck.VBEditor;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Parsing.Symbols
{
    public sealed class ProjectDeclaration : Declaration, ICollection<IInspectionResult>
    {
        private readonly List<ProjectReference> _projectReferences;

        public ProjectDeclaration(
            QualifiedMemberName qualifiedName,
            string name,
            bool isBuiltIn,
            IVBProject project)
            : base(
                  qualifiedName,
                  null,
                  (Declaration)null,
                  name,
                  null,
                  false,
                  false,
                  Accessibility.Implicit,
                  DeclarationType.Project,
                  null,
                  Selection.Home,
                  false,
                  null,
                  isBuiltIn)
        {
            _project = project;
            _projectReferences = new List<ProjectReference>();
        }

        public ProjectDeclaration(ComProject project, QualifiedModuleName module)
            : this(module.QualifyMemberName(project.Name), project.Name, true, null)
        {
            MajorVersion = project.MajorVersion;
            MinorVersion = project.MinorVersion;
        }

        public long MajorVersion { get; set; }
        public long MinorVersion { get; set; }

        public IReadOnlyList<ProjectReference> ProjectReferences
        {
            get
            {
                return _projectReferences.OrderBy(reference => reference.Priority).ToList();
            }
        }

        private readonly IVBProject _project;
        /// <summary>
        /// Gets a reference to the VBProject the declaration is made in.
        /// </summary>
        /// <remarks>
        /// This property is intended to differenciate identically-named VBProjects.
        /// </remarks>
        public override IVBProject Project { get { return _project; } }

        public void AddProjectReference(string referencedProjectId, int priority)
        {
            if (_projectReferences.Any(p => p.ReferencedProjectId == referencedProjectId))
            {
                return;
            }
            _projectReferences.Add(new ProjectReference(referencedProjectId, priority));
        }

        private string _displayName;
        /// <summary>
        /// WARNING: This property has side effects. It changes the ActiveVBProject, which causes a flicker in the VBE.
        /// This should only be called if it is *absolutely* necessary.
        /// </summary>
        public override string ProjectDisplayName
        {
            get
            {
                if (_displayName != null)
                {
                    return _displayName;
                }
                _displayName = _project != null ? _project.ProjectDisplayName : string.Empty;
                return _displayName;
            }
        }

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
