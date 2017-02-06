namespace Rubberduck.Parsing.VBA
{
    //note: ordering of the members is important
    public enum ParserState
    {
        /// <summary>
        /// Parse was requested but hasn't started yet.
        /// </summary>
        Pending,
        /// <summary>
        /// Project references are being loaded into parser state.
        /// </summary>
        LoadingReference,
        /// <summary>
        /// Code from modified modules is being parsed.
        /// </summary>
        Parsing,
        /// <summary>
        /// Parse tree is waiting to be walked for identifier resolution.
        /// </summary>
        Parsed,
        /// <summary>
        /// Resolving declarations.
        /// </summary>
        ResolvingDeclarations,
        /// <summary>
        /// Resolved declarations.
        /// </summary>
        ResolvedDeclarations,
        /// <summary>
        /// Resolving identifier references.
        /// </summary>
        ResolvingReferences,
        /// <summary>
        /// Parser state is in sync with the actual code in the VBE.
        /// </summary>
        Resolved,
        /// <summary>
        /// Parsing is completed, running code inspections.
        /// </summary>
        Inspecting,
        /// <summary>
        /// Parser state and inspection results are in sync with the actual code in the VBE.
        /// </summary>
        Ready,
        /// <summary>
        /// Parsing could not be completed for one or more modules.
        /// </summary>
        Error,
        /// <summary>
        /// Parsing completed, but identifier references could not be resolved for one or more modules.
        /// </summary>
        ResolverError,
        /// <summary>
        /// This component doesn't need a state.  Use for built-in declarations.
        /// </summary>
        None,
    }

    public static class ParserStateExtensions
    {
        /// <summary>
        /// Returns true if resolver succeeded and inspections are running or have completed.
        /// </summary>
        /// <param name="state"></param>
        /// <returns></returns>
        public static bool IsResolvedOrReady(this ParserState state)
        {
            return state == ParserState.Resolved
                || state == ParserState.Ready;
        }
    }
}
