using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
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

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => item.IsBuiltIn && item.IdentifierName == "Excel");
            if (excel == null) { return Enumerable.Empty<IInspectionResult>(); }

            var members = new HashSet<string>(BuiltInDeclarations.Where(decl => decl.DeclarationType == DeclarationType.Function &&
                                                                        decl.ParentDeclaration != null && 
                                                                        decl.ParentDeclaration.ComponentName.Equals("WorksheetFunction"))
                                                                 .Select(decl => decl.IdentifierName));

            var usages = BuiltInDeclarations.Where(decl => decl.References.Any() &&
                                                           decl.ProjectName.Equals("Excel") &&
                                                           decl.ComponentName.Equals("Application") &&
                                                           members.Contains(decl.IdentifierName));

            return (from usage in usages
                from reference in usage.References.Where(use => !IsIgnoringInspectionResultFor(use, AnnotationName))
                let module = reference.ParentScoping.ParentDeclaration
                select new ApplicationWorksheetFunctionInspectionResult(this, new InspectionResultTarget(module, reference), usage.IdentifierName));
        }

        public override void Execute()
        {
            var excel = State.DeclarationFinder.Projects.SingleOrDefault(item => item.IsBuiltIn && item.IdentifierName == "Excel");
            if (excel == null) { return ; }

            var members = new HashSet<string>(BuiltInDeclarations.Where(decl => decl.DeclarationType == DeclarationType.Function &&
                                                                        decl.ParentDeclaration != null &&
                                                                        decl.ParentDeclaration.ComponentName.Equals("WorksheetFunction"))
                                                                 .Select(decl => decl.IdentifierName));

            var usages = BuiltInDeclarations.Where(decl => decl.References.Any() &&
                                                           decl.ProjectName.Equals("Excel") &&
                                                           decl.ComponentName.Equals("Application") &&
                                                           members.Contains(decl.IdentifierName));

            var issues = (
                from usage in usages
                from reference in usage.References.Where(use => !IsIgnoringInspectionResultFor(use, AnnotationName))
                select reference);

            foreach (var issue in issues)
            {
                var module = issue.ParentScoping.ParentDeclaration;
                issue.Annotate(new ApplicationWorksheetFunctionInspectionResult(this, new InspectionResultTarget(module, issue), issue.IdentifierName));
            }
        }
    }
}
