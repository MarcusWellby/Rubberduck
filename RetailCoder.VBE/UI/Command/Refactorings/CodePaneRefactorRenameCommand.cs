using System.Runtime.InteropServices;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.Refactorings.Rename;
using Rubberduck.Settings;
using Rubberduck.UI.Refactorings;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.UI.Command.Refactorings
{
    [ComVisible(false)]
    public class CodePaneRefactorRenameCommand : RefactorCommandBase
    {
        private readonly RubberduckParserState _state;
        private readonly IMessageBox _messageBox;

        public CodePaneRefactorRenameCommand(IVBE vbe, RubberduckParserState state, IMessageBox messageBox) 
            : base (vbe)
        {
            _state = state;
            _messageBox = messageBox;
        }

        public override RubberduckHotkey Hotkey
        {
            get { return RubberduckHotkey.RefactorRename; }
        }

        protected override bool CanExecuteImpl(object parameter)
        {
            if (!_state.Status.IsResolvedOrReady() || Vbe.ActiveCodePane == null)
            {
                return false;
            }

            var target = _state.FindSelectedDeclaration(Vbe.ActiveCodePane);
            return target != null && !target.IsBuiltIn;
        }

        protected override void ExecuteImpl(object parameter)
        {
            if (Vbe.ActiveCodePane == null) { return; }

            Declaration target;
            if (parameter != null)
            {
                target = parameter as Declaration;
            }
            else
            {
                target = _state.FindSelectedDeclaration(Vbe.ActiveCodePane);
            }

            if (target == null || target.IsBuiltIn)
            {
                return;
            }

            using (var view = new RenameDialog())
            {
                var factory = new RenamePresenterFactory(Vbe, view, _state, _messageBox);
                var refactoring = new RenameRefactoring(Vbe, factory, _messageBox, _state);

                refactoring.Refactor(target);
            }
        }
    }
}
