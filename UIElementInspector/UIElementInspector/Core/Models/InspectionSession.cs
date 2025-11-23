using System;
using System.Collections.Generic;

namespace UIElementInspector.Core.Models
{
    /// <summary>
    /// Represents a complete inspection session with all collected data
    /// </summary>
    public class InspectionSession
    {
        public Guid SessionId { get; set; } = Guid.NewGuid();
        public string SessionName { get; set; } = "New Session";
        public DateTime CreatedDate { get; set; } = DateTime.Now;
        public DateTime LastModifiedDate { get; set; } = DateTime.Now;
        public string Description { get; set; } = string.Empty;

        // Collection Settings
        public CollectionProfile CollectionProfile { get; set; } = CollectionProfile.Standard;
        public string DetectionMethod { get; set; } = string.Empty;

        // Collected Data
        public List<ElementInfo> CollectedElements { get; set; } = new List<ElementInfo>();
        public Dictionary<string, byte[]> Screenshots { get; set; } = new Dictionary<string, byte[]>();
        public Dictionary<string, string> SourceCodes { get; set; } = new Dictionary<string, string>();

        // Session Metadata
        public string ApplicationName { get; set; } = string.Empty;
        public string ApplicationVersion { get; set; } = string.Empty;
        public string OperatingSystem { get; set; } = Environment.OSVersion.ToString();
        public string MachineName { get; set; } = Environment.MachineName;
        public string UserName { get; set; } = Environment.UserName;

        // Statistics
        public int TotalElementsCollected => CollectedElements?.Count ?? 0;
        public int TotalScreenshots => Screenshots?.Count ?? 0;
        public int TotalSourceCodes => SourceCodes?.Count ?? 0;
        public TimeSpan TotalCollectionTime { get; set; }

        // Tags and Notes
        public List<string> Tags { get; set; } = new List<string>();
        public string Notes { get; set; } = string.Empty;

        /// <summary>
        /// Gets a summary of the session
        /// </summary>
        public string GetSummary()
        {
            return $"Session: {SessionName}\n" +
                   $"Created: {CreatedDate:yyyy-MM-dd HH:mm:ss}\n" +
                   $"Elements: {TotalElementsCollected}\n" +
                   $"Screenshots: {TotalScreenshots}\n" +
                   $"Source Codes: {TotalSourceCodes}\n" +
                   $"Profile: {CollectionProfile}\n" +
                   $"Detection: {DetectionMethod}";
        }

        /// <summary>
        /// Validates the session data
        /// </summary>
        public List<string> Validate()
        {
            var errors = new List<string>();

            if (string.IsNullOrWhiteSpace(SessionName))
                errors.Add("Session name cannot be empty");

            if (CollectedElements == null)
                errors.Add("Collected elements list cannot be null");

            if (Screenshots == null)
                errors.Add("Screenshots dictionary cannot be null");

            if (SourceCodes == null)
                errors.Add("Source codes dictionary cannot be null");

            return errors;
        }
    }
}
