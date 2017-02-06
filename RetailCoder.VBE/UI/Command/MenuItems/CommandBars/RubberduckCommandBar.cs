using System;
using System.Collections.Generic;
using Rubberduck.Parsing;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.Parsing.Symbols;
using Rubberduck.Parsing.VBA;
using Rubberduck.VBEditor.SafeComWrappers.Abstract;
using Rubberduck.VBEditor.SafeComWrappers.Office.Core;

namespace Rubberduck.UI.Command.MenuItems.CommandBars
{
    public class RubberduckCommandBar : AppCommandBarBase, IDisposable
    {
        private readonly IContextFormatter _formatter;

        public RubberduckCommandBar(IEnumerable<ICommandMenuItem> items, IContextFormatter formatter) 
            : base("Rubberduck", CommandBarPosition.Top, items)
        {
            _formatter = formatter;
        }

        public void SetStatusLabelCaption(ParserState state, int? errorCount = null)
        {
            var caption = ParsingText.ResourceManager.GetString("ParserState_" + state, Settings.Settings.Culture);
            SetStatusLabelCaption(caption, errorCount);
        }

        public void SetStatusLabelCaption(string caption, int? errorCount = null)
        {
            var reparseCommandButton = FindChildByTag(typeof(ReparseCommandMenuItem).FullName) as ReparseCommandMenuItem;
            if (reparseCommandButton == null) { return; }

            var showErrorsCommandButton = FindChildByTag(typeof(ShowParserErrorsCommandMenuItem).FullName) as ShowParserErrorsCommandMenuItem;
            if (showErrorsCommandButton == null) { return; }

            UiDispatcher.Invoke(() =>
            {
                reparseCommandButton.SetCaption(caption);
                reparseCommandButton.SetToolTip(string.Format(RubberduckUI.ReparseToolTipText, caption));
                if (errorCount.HasValue && errorCount.Value > 0)
                {
                    showErrorsCommandButton.SetToolTip(string.Format(RubberduckUI.ParserErrorToolTipText, errorCount.Value));
                }
            });
            Localize();
        }

        public string GetContextSelectionCaption(ICodePane activeCodePane, Declaration declaration)
        {
            return _formatter.Format(activeCodePane, declaration);
        }

        public void SetContextSelectionCaption(string caption, int contextReferenceCount, IEnumerable<IInspectionResult> inspectionResults)
        {
            var contextLabel = FindChildByTag(typeof(ContextSelectionLabelMenuItem).FullName) as ContextSelectionLabelMenuItem;
            if (contextLabel == null) { return; }

            var contextReferences = FindChildByTag(typeof(ReferenceCounterLabelMenuItem).FullName) as ReferenceCounterLabelMenuItem;
            if (contextReferences == null) { return; }

            var issues = FindChildByTag(typeof(InspectionResultsLabelMenuItem).FullName) as InspectionResultsLabelMenuItem;

            UiDispatcher.Invoke(() =>
            {
                contextLabel.SetCaption(caption);
                contextReferences.SetCaption(contextReferenceCount);
                if (issues != null)
                {
                    issues.SetCaption(inspectionResults);
                }
            });
            Localize();
        }

        public void Dispose()
        {
            //note: doing this wrecks the teardown process. counter-intuitive? sure. but hey it works.
            //RemoveChildren();
            //Item.Delete();
            //Item.Release(true);
        }
    }

    public enum RubberduckCommandBarItemDisplayOrder
    {
        RequestReparse,
        ShowErrors,
        ContextStatus,
        ContextRefCount,
        InspectionResults
    }
}