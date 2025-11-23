namespace UIElementInspector.Core.Models
{
    /// <summary>
    /// Defines the depth of data collection
    /// </summary>
    public enum CollectionProfile
    {
        /// <summary>
        /// Only basic properties: ID, Name, Type, Position
        /// </summary>
        Quick,

        /// <summary>
        /// Standard properties: Basic + Common attributes
        /// </summary>
        Standard,

        /// <summary>
        /// All available properties from all technologies
        /// </summary>
        Full,

        /// <summary>
        /// User-defined custom selection of properties
        /// </summary>
        Custom
    }

    /// <summary>
    /// Defines which properties to collect in Custom mode
    /// </summary>
    public class CustomCollectionSettings
    {
        public bool IncludeBasicProperties { get; set; } = true;
        public bool IncludeUIAutomation { get; set; } = true;
        public bool IncludeWebProperties { get; set; } = true;
        public bool IncludeHierarchy { get; set; } = true;
        public bool IncludeSelectors { get; set; } = true;
        public bool IncludeAccessibility { get; set; } = true;
        public bool IncludeStyles { get; set; } = false;
        public bool IncludeScreenshot { get; set; } = false;
        public bool IncludeSourceCode { get; set; } = false;
        public bool IncludeEventListeners { get; set; } = false;
    }
}