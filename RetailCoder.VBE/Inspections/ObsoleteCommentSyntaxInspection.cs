using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class ObsoleteCommentSyntaxInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _results;

        public ObsoleteCommentSyntaxInspection(RubberduckParserState state) : base(state, CodeInspectionSeverity.Suggestion) { }

        public override string Meta { get { return InspectionsUI.ObsoleteCommentSyntaxInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteCommentSyntaxInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public override void Execute()
        {
            if (ParseTreeResults == null) { return; }

            var issues = from issue in ParseTreeResults
                where !IsIgnoringInspectionResultFor(issue.ModuleName.Component, issue.Context.Start.Line)
                let target = issue.Context
                let module = issue.ModuleName.Component.CodeModule
                select new
                {
                    Target = target,
                    Result = new ObsoleteCommentSyntaxInspectionResult(this, target.Target, module)
                };

            foreach (var issue in issues)
            {
                issue.Target.Add(issue.Result);
            }
        }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _results = results;
        }

        private IEnumerable<QualifiedContext<VBAParser.RemCommentContext>> ParseTreeResults { get { return _results.OfType<QualifiedContext<VBAParser.RemCommentContext>>(); } }


        public class ObsoleteCommentSyntaxListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.RemCommentContext> _contexts = new List<VBAParser.RemCommentContext>();

            public IEnumerable<VBAParser.RemCommentContext> Contexts
            {
                get { return _contexts; }
            }

            public override void ExitRemComment(VBAParser.RemCommentContext context)
            {
                _contexts.Add(context);
            }
        }
    }
}
