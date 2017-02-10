using System;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.Results
{
    public class ObsoleteCommentSyntaxInspectionResult : InspectionResultBase
    {
        private readonly ICodeModule _module;
        private IEnumerable<QuickFixBase> _quickFixes;

        public ObsoleteCommentSyntaxInspectionResult(IInspection inspection, IInspectionResultTarget target, ICodeModule module)
            : base(inspection, target, InspectionsUI.ObsoleteCommentSyntaxInspectionResultFormat)
        {
            _module = module;
        }

        [Obsolete]
        public ObsoleteCommentSyntaxInspectionResult(IInspection inspection, QualifiedContext<ParserRuleContext> qualifiedContext)
            : base(inspection, qualifiedContext.ModuleName, qualifiedContext.Context)
        { }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new ReplaceObsoleteCommentMarkerQuickFix(Context, QualifiedSelection),
                    new RemoveCommentQuickFix(Context, QualifiedSelection), 
                    new IgnoreOnceQuickFix(_module, Target, Inspection.AnnotationName)
                });
            }
        }
    }
}
