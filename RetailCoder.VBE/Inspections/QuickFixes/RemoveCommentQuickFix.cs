using Antlr4.Runtime;
using Rubberduck.Inspections.Abstract;
using Rubberduck.Inspections.Resources;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor;

namespace Rubberduck.Inspections.QuickFixes
{
    public class RemoveCommentQuickFix : QuickFixBase
    {
        public RemoveCommentQuickFix(ParserRuleContext context, QualifiedSelection selection)
            : base(context, selection, InspectionsUI.RemoveObsoleteStatementQuickFix)
        { }

        public override void Fix()
        {
            var module = QualifiedSelection.QualifiedName.Component.CodeModule;

            if (module.IsWrappingNullReference)
            {
                return;                
            }

            var start = Context.Start.Line;
            var commentLine = module.GetLines(start, QualifiedSelection.Selection.LineCount);
            var newLine = commentLine.Substring(0, Context.Start.Column).TrimEnd();

            module.DeleteLines(start, QualifiedSelection.Selection.LineCount);
            if (newLine.TrimStart().Length > 0)
            {
                module.InsertLines(start, newLine);
            }
        }
    }
}