using System;
using System.Windows.Threading;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;

namespace Rubberduck.Inspections.Abstract
{
    public abstract class QuickFixBase : IQuickFix
    {
        private readonly ParserRuleContext _context;
        private readonly QualifiedSelection _qualifiedSelection;

        private readonly ICodeModule _module;
        private readonly IInspectionResultTarget _target;
        private readonly string _description;

        protected QuickFixBase(ICodeModule module, IInspectionResultTarget target, string description)
        {
            Dispatcher.CurrentDispatcher.Thread.CurrentCulture = UI.Settings.Settings.Culture;
            Dispatcher.CurrentDispatcher.Thread.CurrentUICulture = UI.Settings.Settings.Culture;

            _module = module;
            _target = target;
            _description = description;
        }

        [Obsolete]
        protected QuickFixBase(ParserRuleContext context, QualifiedSelection selection, string description)
        {
            Dispatcher.CurrentDispatcher.Thread.CurrentCulture = UI.Settings.Settings.Culture;
            Dispatcher.CurrentDispatcher.Thread.CurrentUICulture = UI.Settings.Settings.Culture;

            _context = context;
            _qualifiedSelection = selection;
            _description = description;
        }

        public string Description { get { return _description; } }

        protected ParserRuleContext Context { get { return _context; } }

        public QualifiedSelection QualifiedSelection { get { return _qualifiedSelection; } }

        public bool IsCancelled { get; set; }

        protected abstract void Fix(ICodeModule module, IInspectionResultTarget target);

        public virtual void Fix()
        {
            Fix(_module, _target);
        }

        /// <summary>
        /// Indicates whether this quickfix can be applied to all inspection results in module.
        /// </summary>
        /// <remarks>
        /// If both <see cref="CanFixInModule"/> and <see cref="CanFixInProject"/> are set to <c>false</c>,
        /// then the quickfix is only applicable to the current/selected inspection result.
        /// </remarks>
        public virtual bool CanFixInModule { get { return true; } }

        /// <summary>
        /// Indicates whether this quickfix can be applied to all inspection results in project.
        /// </summary>
        /// <remarks>
        /// If both <see cref="CanFixInModule"/> and <see cref="CanFixInProject"/> are set to <c>false</c>,
        /// then the quickfix is only applicable to the current/selected inspection result.
        /// </remarks>
        public virtual bool CanFixInProject { get { return true; } }
    }
}
