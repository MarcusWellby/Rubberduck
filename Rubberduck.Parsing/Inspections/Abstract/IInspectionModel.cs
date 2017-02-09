namespace Rubberduck.Parsing.Inspections.Abstract
{
    /// <summary>
    /// An interface that abstracts the data structure for a code inspection
    /// </summary>
    public interface IInspectionModel
    {
        /// <summary>
        /// Gets the inspection name.
        /// </summary>
        string Name { get; }

        /// <summary>
        /// Gets a string that contains additional/meta information about an inspection.
        /// </summary>
        string Meta { get; }
        
        string AnnotationName { get; }

        /// <summary>
        /// Gets a short description for the code inspection.
        /// </summary>
        string Description { get; }

        /// <summary>
        /// Gets a value indicating the type of the code inspection.
        /// </summary>
        CodeInspectionType InspectionType { get; }

        /// <summary>
        /// Gets a value indicating the default severity level of the code inspection.
        /// </summary>
        CodeInspectionSeverity DefaultSeverity { get; }

        /// <summary>
        /// Gets a value indicating the severity level of the code inspection.
        /// </summary>
        CodeInspectionSeverity Severity { get; set; }
    }
}
