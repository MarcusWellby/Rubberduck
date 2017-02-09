using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using Antlr4.Runtime;
using Rubberduck.Parsing.Grammar;
using Rubberduck.Parsing.Inspections.Abstract;
using Rubberduck.VBEditor;

namespace Rubberduck.Parsing
{
    public static class ParserRuleContextExtensions
    {
        /// <summary>
        /// Allows annotating a context with inspection results, iterating them, and clearing them.
        /// </summary>
        /// <returns>The context instance, cast to an <see cref="ICollection{IInspectionResult}"/>.</returns>
        /// <exception cref="InvalidCastException">Thrown when context does not implement <see cref="ICollection{IInspectionResult}"/>.</exception>
        /// <remarks>Backed with a <see cref="ConcurrentBag{IInspectionResult}"/>, the collection does not support <see cref="ICollection{IInspectionResult}.Remove"/>.</remarks>
        public static ICollection<IInspectionResult> InspectionResults(this ParserRuleContext context)
        {
            return (ICollection<IInspectionResult>) context;
        }

        public static Selection GetSelection(this ParserRuleContext context)
        {
            if (context == null)
                return Selection.Home;

            // ANTLR indexes are 0-based, but VBE's are 1-based.
            // 1 is the default value that will select all lines. Replace zeroes with ones.
            // See also: https://msdn.microsoft.com/en-us/library/aa443952(v=vs.60).aspx

            return new Selection(context.Start.Line == 0 ? 1 : context.Start.Line,
                                 context.Start.Column + 1,
                                 context.Stop.Line == 0 ? 1 : context.Stop.Line,
                                 context.Stop.Column + context.Stop.Text.Length + 1);
        }

        //This set of overloads returns the selection for the entire procedure statement body, i.e. Public Function Foo(bar As String) As String
        public static Selection GetProcedureSelection(this VBAParser.FunctionStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.SubStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertyGetStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertyLetStmtContext context) { return GetProcedureContextSelection(context); }
        public static Selection GetProcedureSelection(this VBAParser.PropertySetStmtContext context) { return GetProcedureContextSelection(context); }

        private static Selection GetProcedureContextSelection(ParserRuleContext context)
        {
            var endContext = context.GetRuleContext<VBAParser.EndOfStatementContext>(0);
            return new Selection(context.Start.Line == 0 ? 1 : context.Start.Line,
                                 context.Start.Column + 1,
                                 endContext.Start.Line == 0 ? 1 : endContext.Start.Line,
                                 endContext.Start.Column + 1);
        }
    }
}
