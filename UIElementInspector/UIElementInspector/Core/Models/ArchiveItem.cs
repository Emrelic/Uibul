using System;
using System.Collections.Generic;

namespace UIElementInspector.Core.Models
{
    /// <summary>
    /// Represents an archived capture item
    /// </summary>
    public class ArchiveItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString();
        public string Name { get; set; } = string.Empty;
        public string FolderPath { get; set; } = string.Empty;
        public DateTime CaptureTime { get; set; } = DateTime.Now;
        public string CaptureType { get; set; } = "FullCapture"; // FullCapture, QuickExport
        public int FileCount { get; set; }
        public List<string> FilePaths { get; set; } = new List<string>();
        public string Notes { get; set; } = string.Empty;

        // Element info for quick reference
        public string ElementName { get; set; } = string.Empty;
        public string ElementType { get; set; } = string.Empty;
        public string WindowTitle { get; set; } = string.Empty;
    }

    /// <summary>
    /// Archive index for managing all archived items
    /// </summary>
    public class ArchiveIndex
    {
        public List<ArchiveItem> Items { get; set; } = new List<ArchiveItem>();
        public DateTime LastUpdated { get; set; } = DateTime.Now;
        public string Version { get; set; } = "1.0";
    }
}
