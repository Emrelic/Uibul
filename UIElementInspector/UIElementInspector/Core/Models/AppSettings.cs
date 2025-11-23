using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Newtonsoft.Json;

namespace UIElementInspector.Core.Models
{
    /// <summary>
    /// Application settings and user preferences
    /// </summary>
    public class AppSettings
    {
        private static readonly string SettingsFilePath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "UIElementInspector",
            "settings.json");

        // Collection Settings
        public CollectionProfile DefaultCollectionProfile { get; set; } = CollectionProfile.Standard;

        // Export Settings
        public List<string> ExportFormats { get; set; } = new List<string> { "CSV", "JSON", "XML" };
        public string ExportDirectory { get; set; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
            "UI Inspector Exports");
        public bool AutoExportOnCapture { get; set; } = false;
        public bool OrganizeExportsByDate { get; set; } = true;

        // Screenshot Settings
        public string ScreenshotFormat { get; set; } = "PNG";
        public int JpegQuality { get; set; } = 90;
        public bool AutoCaptureScreenshot { get; set; } = true;
        public bool IncludeTimestampInFilename { get; set; } = true;

        // Performance Settings
        public int MouseHoverDelay { get; set; } = 500; // milliseconds
        public int MaxTreeDepth { get; set; } = 20;
        public bool EnableDetectionThrottling { get; set; } = true;
        public bool CacheElements { get; set; } = true;

        // Hotkey Settings (for future implementation)
        public string HotkeyStartStop { get; set; } = "Ctrl+Shift+I";
        public string HotkeyCaptureElement { get; set; } = "Ctrl+Click";
        public string HotkeyRegionSelection { get; set; } = "Ctrl+Shift+R";
        public string HotkeyExportAll { get; set; } = "Ctrl+Shift+E";

        // UI Settings
        public bool AlwaysOnTop { get; set; } = false;
        public bool MinimizeToTray { get; set; } = false;
        public bool ShowNotifications { get; set; } = true;
        public bool ShowTooltips { get; set; } = true;

        // Advanced Settings
        public bool EnableLogging { get; set; } = true;
        public string LogLevel { get; set; } = "Info"; // Debug, Info, Warning, Error
        public int MaxLogFileSizeMB { get; set; } = 10;

        /// <summary>
        /// Loads settings from disk or creates default settings
        /// </summary>
        public static AppSettings Load()
        {
            try
            {
                if (File.Exists(SettingsFilePath))
                {
                    var json = File.ReadAllText(SettingsFilePath);
                    var settings = JsonConvert.DeserializeObject<AppSettings>(json);
                    return settings ?? CreateDefault();
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading settings: {ex.Message}");
            }

            return CreateDefault();
        }

        /// <summary>
        /// Saves settings to disk
        /// </summary>
        public void Save()
        {
            try
            {
                var directory = Path.GetDirectoryName(SettingsFilePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                var json = JsonConvert.SerializeObject(this, Formatting.Indented);
                File.WriteAllText(SettingsFilePath, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error saving settings: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates default settings
        /// </summary>
        public static AppSettings CreateDefault()
        {
            return new AppSettings();
        }

        /// <summary>
        /// Validates settings and returns validation errors
        /// </summary>
        public List<string> Validate()
        {
            var errors = new List<string>();

            if (string.IsNullOrWhiteSpace(ExportDirectory))
            {
                errors.Add("Export directory cannot be empty");
            }

            if (!ExportFormats.Any())
            {
                errors.Add("At least one export format must be selected");
            }

            if (MouseHoverDelay < 100 || MouseHoverDelay > 5000)
            {
                errors.Add("Mouse hover delay must be between 100 and 5000 milliseconds");
            }

            if (MaxTreeDepth < 1 || MaxTreeDepth > 100)
            {
                errors.Add("Max tree depth must be between 1 and 100");
            }

            if (JpegQuality < 1 || JpegQuality > 100)
            {
                errors.Add("JPEG quality must be between 1 and 100");
            }

            return errors;
        }

        /// <summary>
        /// Applies settings to the application
        /// </summary>
        public void Apply()
        {
            // This method will be called by MainWindow to apply settings
            // Implementation will vary based on what needs to be updated

            // Create export directory if it doesn't exist
            if (!Directory.Exists(ExportDirectory))
            {
                try
                {
                    Directory.CreateDirectory(ExportDirectory);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error creating export directory: {ex.Message}");
                }
            }
        }
    }
}
