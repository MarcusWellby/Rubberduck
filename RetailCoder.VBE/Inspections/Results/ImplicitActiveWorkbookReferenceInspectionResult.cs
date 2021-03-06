using System;
using System.Collections.Generic;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections.Results
{
    public class ImplicitActiveWorkbookReferenceInspectionResult : InspectionResultBase
    {
        private readonly IdentifierReference _reference;
        private IEnumerable<QuickFixBase> _quickFixes;

        public ImplicitActiveWorkbookReferenceInspectionResult(IInspection inspection, InspectionResultTarget target, string name)
            : base(inspection, name) { }

        [Obsolete]
        public ImplicitActiveWorkbookReferenceInspectionResult(IInspection inspection, IdentifierReference reference)
            : base(inspection, reference.QualifiedModuleName, reference.Context)
        {
            _reference = reference;
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new IgnoreOnceQuickFix(_reference.Context, QualifiedSelection, Inspection.AnnotationName)
                });
            }
        }

        public override string Description
        {
            get { return string.Format(InspectionsUI.ImplicitActiveSheetReferenceInspectionResultFormat, Context.GetText() /*_reference.Declaration.IdentifierName*/); }
        }
    }
}