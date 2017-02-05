using Antlr4.Runtime;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IQuickFix
    {
        string Description { get; }
        ParserRuleContext Context { get; }
        QualifiedSelection Selection { get; }
        bool IsCancelled { get; set; }
        void Fix();
        bool CanFixInModule { get; }
        bool CanFixInProject { get; }
    }
}
