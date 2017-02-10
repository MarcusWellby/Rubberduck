using Rubberduck.VBEditor;

namespace Rubberduck.Parsing.Inspections.Abstract
{
    public interface IQuickFix
    {
        string Description { get; }

        void Fix();
        bool CanFixInModule { get; }
        bool CanFixInProject { get; }

        QualifiedSelection QualifiedSelection { get; }

        bool IsCancelled { get; set; }
    }
}
