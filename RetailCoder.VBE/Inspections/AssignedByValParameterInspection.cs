using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Inspections.Results;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;

namespace Rubberduck.Inspections
{
    public sealed class AssignedByValParameterInspection : InspectionBase
    {
        public AssignedByValParameterInspection(RubberduckParserState state)
            : base(state)
        {
            Severity = DefaultSeverity;
        }

        public override string Meta { get { return InspectionsUI.AssignedByValParameterInspectionMeta; } }
        public override string Description { get { return InspectionsUI.AssignedByValParameterInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            var parameters = State.DeclarationFinder.UserDeclarations(DeclarationType.Parameter)
                .OfType<ParameterDeclaration>()
                .Where(item => !item.IsByRef 
                    && !IsIgnoringInspectionResultFor(item, AnnotationName)
                    && item.References.Any(reference => reference.IsAssignment))
                .ToList();

            return parameters
                .Select(param => new AssignedByValParameterInspectionResult(this, param))
                .ToList();
        }
    }
}
