using System;

namespace Rubberduck.Parsing.VBA
{
    public class ParseRequestEventArgs : EventArgs
    {
        private readonly object _requestor;
        private readonly bool _runInspections;

        /// <summary>
        /// Encapsulates the arguments of a parse request.
        /// </summary>
        /// <param name="requestor">The object requesting the parse.</param>
        /// <param name="runInspections">True if inspections should run on a successful parse.</param>
        public ParseRequestEventArgs(object requestor, bool runInspections)
        {
            _requestor = requestor;
            _runInspections = runInspections;
        }

        public object Requestor { get { return _requestor; } }

        public bool RunInspections { get { return _runInspections; } }
    }
}