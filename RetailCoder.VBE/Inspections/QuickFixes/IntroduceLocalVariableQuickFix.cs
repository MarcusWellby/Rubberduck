using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.QuickFixes
{
    public class IntroduceLocalVariableQuickFix : QuickFixBase
    {
        private readonly Declaration _undeclared;

        public IntroduceLocalVariableQuickFix(Declaration undeclared) 
            : base(undeclared.Context, undeclared.QualifiedSelection, InspectionsUI.IntroduceLocalVariableQuickFix)
        {
            _undeclared = undeclared;
        }

        public override bool CanFixInModule { get { return true; } }
        public override bool CanFixInProject { get { return true; } }

        protected override void Fix(ICodeModule module, IInspectionResultTarget target)
        {
            throw new System.NotImplementedException();
        }

        public override void Fix()
        {
            var instruction = Tokens.Dim + ' ' + _undeclared.IdentifierName + ' ' + Tokens.As + ' ' + Tokens.Variant;
            QualifiedSelection.QualifiedName.Component.CodeModule.InsertLines(QualifiedSelection.Selection.StartLine, instruction);
        }
    }
}