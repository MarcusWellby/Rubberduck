using System;
using System.Collections.Generic;
using System.Linq;
using Rubberduck.Common;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.Results
{
    public class ProcedureCanBeWrittenAsFunctionInspectionResult : InspectionResultBase
    {
        private readonly ICodeModule _module;
        private IEnumerable<QuickFixBase> _quickFixes;
        private readonly QualifiedContext<VBAParser.ArgListContext> _argListQualifiedContext;
        private readonly QualifiedContext<VBAParser.SubStmtContext> _subStmtQualifiedContext;
        private readonly RubberduckParserState _state;

        public ProcedureCanBeWrittenAsFunctionInspectionResult(IInspection inspection, IInspectionResultTarget target, ICodeModule module, string name)
            : base(inspection, target, name)
        {
            _module = module;
        }

        [Obsolete]
        public ProcedureCanBeWrittenAsFunctionInspectionResult(IInspection inspection, RubberduckParserState state,
            QualifiedContext<VBAParser.ArgListContext> argListQualifiedContext,
            QualifiedContext<VBAParser.SubStmtContext> subStmtQualifiedContext)
            : base(inspection, subStmtQualifiedContext.ModuleName, subStmtQualifiedContext.Context.subroutineName())
        {
            _target =
                state.AllUserDeclarations.Single(declaration => declaration.DeclarationType == DeclarationType.Procedure
                                                                &&
                                                                declaration.Context == subStmtQualifiedContext.Context);

            _argListQualifiedContext = argListQualifiedContext;
            _subStmtQualifiedContext = subStmtQualifiedContext;
            _state = state;
        }

        public override IEnumerable<IQuickFix> QuickFixes
        {
            get
            {
                return _quickFixes ?? (_quickFixes = new QuickFixBase[]
                {
                    new ChangeProcedureToFunction(_state, _argListQualifiedContext, _subStmtQualifiedContext, QualifiedSelection),
                    new IgnoreOnceQuickFix(_module, Target, Inspection.AnnotationName)
                });
            }
        }

        private readonly Declaration _target;
    }
}
