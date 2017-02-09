using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using InspectionBase = Rubberduck.Inspections.Abstract.InspectionBase;

namespace Rubberduck.Inspections
{
    public class ApplicationWorksheetFunctionInspection : InspectionBase
    {
        public ApplicationWorksheetFunctionInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Suggestion)
        { }

        public override string Meta { get { return InspectionsUI.ApplicationWorksheetFunctionInspectionMeta; } }
        public override string Description { get { return InspectionsUI.ApplicationWorksheetFunctionInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override void Execute()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => item.IsBuiltIn && item.IdentifierName == "Excel");
            if (excel == null) { return ; }

            var members = new HashSet<string>(BuiltInDeclarations
                .Where(declaration => declaration.DeclarationType == DeclarationType.Function 
                    && declaration.ParentDeclaration != null 
                    && declaration.ParentDeclaration.ComponentName.Equals("WorksheetFunction"))
                .Select(declaration => declaration.IdentifierName));

            var declarations = BuiltInDeclarations
                .Where(declaration => declaration.References.Any() 
                    && declaration.ProjectName.Equals("Excel") 
                    && declaration.ComponentName.Equals("Application") 
                    && members.Contains(declaration.IdentifierName));

            var issues =
                from declaration in declarations
                from reference in declaration.References.Where(use => !IsIgnoringInspectionResultFor(use, AnnotationName))
                let module = reference.ParentScoping.ParentDeclaration
                select new
                {
                    Target = reference.Context,
                    Result = new ApplicationWorksheetFunctionInspectionResult(this,
                            string.Format(InspectionsUI.ApplicationWorksheetFunctionInspectionResultFormat, declaration.IdentifierName))
                };

            foreach (var issue in issues)
            {
                issue.Target.InspectionResults().Add(issue.Result);
            }
        }
    }
}
