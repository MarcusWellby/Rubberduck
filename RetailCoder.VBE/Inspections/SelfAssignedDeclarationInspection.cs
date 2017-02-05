﻿using System.Collections.Generic;
using System.Linq;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.QuickFixes;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.VBA;
using Rubberduck.Parsing.Symbols;

namespace Rubberduck.Inspections
{
    public sealed class SelfAssignedDeclarationInspection : InspectionBase
    {
        public SelfAssignedDeclarationInspection(RubberduckParserState state)
            : base(state, CodeInspectionSeverity.Hint)
        {
        }

        public override string Meta { get { return InspectionsUI.SelfAssignedDeclarationInspectionMeta; } }
        public override string Description { get { return InspectionsUI.SelfAssignedDeclarationInspectionName; } }
        public override CodeInspectionType InspectionType { get { return CodeInspectionType.CodeQualityIssues; } }

        private static readonly IReadOnlyList<string> ValueTypes = new[]
        {
            Tokens.Boolean,
            Tokens.Byte,
            Tokens.Currency,
            Tokens.Date,
            Tokens.Decimal,
            Tokens.Double,
            Tokens.Integer,
            Tokens.Long,
            Tokens.LongLong,
            Tokens.Single,
            Tokens.String
        };

        public override IEnumerable<IInspectionResult> GetInspectionResults()
        {
            return UserDeclarations
                .Where(declaration => declaration.IsSelfAssigned 
                    && declaration.IsTypeSpecified
                    && !ValueTypes.Contains(declaration.AsTypeName)
                    && declaration.DeclarationType == DeclarationType.Variable
                    && (declaration.AsTypeDeclaration == null
                        || declaration.AsTypeDeclaration.DeclarationType != DeclarationType.UserDefinedType)
                    && declaration.ParentScopeDeclaration != null
                    && declaration.ParentScopeDeclaration.DeclarationType.HasFlag(DeclarationType.Member))
                .Select(issue => new SelfAssignedDeclarationInspectionResult(this, issue));
        }
    }
}
