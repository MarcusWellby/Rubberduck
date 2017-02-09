using System;
using System.Collections.Generic;
using System.Linq;
using Antlr4.Runtime;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class InspectionResultBase : IInspectionResult
    {
        private readonly IInspection _inspection;
        private readonly string _description;

        protected InspectionResultBase(IInspection inspection, string description)
        {
            _inspection = inspection;
            _description = description;
        }

        #region obsolete constructors
        [Obsolete]
        protected InspectionResultBase(IInspection inspection, Declaration target)
            : this(inspection, target.QualifiedName.QualifiedModuleName, target.Context)
        {
        }

        /// <summary>
        /// Creates an inspection result.
        /// </summary>
        [Obsolete]
        protected InspectionResultBase(IInspection inspection, QualifiedModuleName qualifiedName, ParserRuleContext context = null)
        {
            _inspection = inspection;
        }

        /// <summary>
        /// Creates an inspection result.
        /// </summary>
        [Obsolete]
        protected InspectionResultBase(IInspection inspection, QualifiedModuleName qualifiedName, ParserRuleContext context, Declaration declaration)
        {
            _inspection = inspection;
        }
        #endregion

        public IInspection Inspection { get { return _inspection; } }

        public string Description { get {return _description; } }

        public virtual IEnumerable<IQuickFix> QuickFixes { get { return Enumerable.Empty<QuickFixBase>(); } }

        public IQuickFix DefaultQuickFix { get { return QuickFixes == null ? null : QuickFixes.FirstOrDefault(); } }

        ///// <summary>
        ///// WARNING: This property can have side effects. It can change the ActiveVBProject if the result has a null Declaration, 
        ///// which causes a flicker in the VBE. This should only be called if it is *absolutely* necessary.
        ///// </summary>
        //public string ToClipboardString()
        //{           
        //    var module = QualifiedSelection.QualifiedName;
        //    var documentName = _target != null ? _target.ProjectDisplayName : string.Empty;
        //    if (string.IsNullOrEmpty(documentName))
        //    {
        //        var component = module.Component;
        //        documentName = component != null ? component.ParentProject.ProjectDisplayName : string.Empty;
        //    }
        //    if (string.IsNullOrEmpty(documentName))
        //    {
        //        documentName = Path.GetFileName(module.ProjectPath);
        //    }

        //    return string.Format(
        //        InspectionsUI.QualifiedSelectionInspection,
        //        Inspection.Severity,
        //        Description,
        //        "(" + documentName + ")",
        //        module.ProjectName,
        //        module.ComponentName,
        //        QualifiedSelection.Selection.StartLine);
        //}

        //public object[] ToArray()
        //{
        //    var module = QualifiedSelection.QualifiedName;
        //    return new object[] { Inspection.Severity.ToString(), module.ProjectName, module.ComponentName, Description, QualifiedSelection.Selection.StartLine, QualifiedSelection.Selection.StartColumn };
        //}
    }
}
