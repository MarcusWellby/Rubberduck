using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Annotations;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public sealed class MissingAnnotationArgumentInspection : InspectionBase, IParseTreeInspection
    {
        private IEnumerable<QualifiedContext> _parseTreeResults;
 
        public MissingAnnotationArgumentInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Error)
        {
        }

        public override string Meta { get { return InspectionsUI.MissingAnnotationArgumentInspectionMeta; } }
        public override string Description { get { return InspectionsUI.MissingAnnotationArgumentInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }
        public IEnumerable<QualifiedContext<VBAParser.AnnotationContext>> ParseTreeResults { get { return _parseTreeResults.OfType<QualifiedContext<VBAParser.AnnotationContext>>(); } }

        public void SetResults(IEnumerable<QualifiedContext> results)
        {
            _parseTreeResults = results;
        }

        public override void Execute()
        {
            if (ParseTreeResults == null) { return; }

            var issues = 
                from result in ParseTreeResults
                let module = State.DeclarationFinder.UserDeclarations(DeclarationType.Module)
                                                    .SingleOrDefault(m => m.QualifiedName == result.MemberName)
                let context = result.Context
                let name = context.annotationName().GetText()
                where name == AnnotationType.Ignore.ToString()
                   || name == AnnotationType.Folder.ToString()
                where context.annotationArgList() == null
                select new
                {
                    Target = context, 
                    Result = new MissingAnnotationArgumentInspectionResult(this, 
                        string.Format(InspectionsUI.MissingAnnotationArgumentInspectionResultFormat, name))
                };

            foreach (var issue in issues)
            {
                issue.Target.InspectionResults().Add(issue.Result);
            }
        }

        public class InvalidAnnotationStatementListener : VBAParserBaseListener
        {
            private readonly IList<VBAParser.AnnotationContext> _contexts = new List<VBAParser.AnnotationContext>();
            public IEnumerable<VBAParser.AnnotationContext> Contexts { get { return _contexts; } }

            public override void ExitAnnotation(VBAParser.AnnotationContext context)
            {
                if (context.annotationName() != null)
                {
                    _contexts.Add(context);
                }
            }
        }
    }
}
