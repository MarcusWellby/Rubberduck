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
    public sealed class ObsoleteCallStatementInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;

        public ObsoleteCallStatementInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        {
        }

        public override string Meta { get { return InspectionsUI.ObsoleteCallStatementInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ObsoleteCallStatementInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.LanguageOpportunities; } }

        public IEnumerable<QualifiedContext<VBAParser.CallStmtContext>> ParseTreeResults { get { return _parseTreeResults.OfType<QualifiedContext<VBAParser.CallStmtContext>>(); } }
        public void SetResults(IEnumerable<QualifiedContext> results) { _parseTreeResults = results; }

        public override void Execute()
        {
            if (ParseTreeResults == null) { return; }

            foreach (var context in ParseTreeResults.Where(context => !IsIgnoringInspectionResultFor(context.ModuleName.Component, context.Context.Start.Line)))
            {
                var module = context.ModuleName.Component.CodeModule;
                {
                    var lines = module.GetLines(context.Context.Start.Line, context.Context.Stop.Line - context.Context.Start.Line + 1);

                    var stringStrippedLines = string.Join(string.Empty, lines).StripStringLiterals();

                    int commentIndex;
                    if (stringStrippedLines.HasComment(out commentIndex))
                    {
                        stringStrippedLines = stringStrippedLines.Remove(commentIndex);
                    }

                    if (!stringStrippedLines.Contains(":"))
                    {
                        var target = context.Context.Target;
                        var result = new ObsoleteCallStatementUsageInspectionResult(this, target, Tokens.Call);
                        target.Add(result);
                    }
                }
            }
        }

        public class ObsoleteCallStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.CallStmtContext> _contexts = new List<VBAParser.CallStmtContext>();
            public IEnumerable<VBAParser.CallStmtContext> Contexts { get { return _contexts; } }

            public override void ExitCallStmt(VBAParser.CallStmtContext context)
            {
                if (context.CALL() != null)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
