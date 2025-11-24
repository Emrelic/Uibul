using System;
using System.Collections.Generic;
using System.Windows;

namespace UIElementInspector.Core.Models
{
    /// <summary>
    /// Universal element information container that holds data from all detection technologies
    /// </summary>
    public class ElementInfo
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public DateTime CaptureTime { get; set; } = DateTime.Now;
        public string DetectionMethod { get; set; } // UI Automation, WebView2, MSHTML, Playwright, etc.

        // === Basic Properties ===
        public string ElementType { get; set; }
        public string Name { get; set; }
        public string ClassName { get; set; }
        public string Value { get; set; }
        public string Description { get; set; }

        // === UI Automation Properties ===
        public string AutomationId { get; set; }
        public string ControlType { get; set; }
        public string LocalizedControlType { get; set; }
        public string FrameworkId { get; set; }
        public int ProcessId { get; set; }
        public string RuntimeId { get; set; }
        public IntPtr NativeWindowHandle { get; set; }
        public string ItemStatus { get; set; }
        public string HelpText { get; set; }
        public string AcceleratorKey { get; set; }
        public string AccessKey { get; set; }
        public bool IsEnabled { get; set; }
        public bool IsOffscreen { get; set; }
        public bool HasKeyboardFocus { get; set; }
        public bool IsKeyboardFocusable { get; set; }
        public bool IsPassword { get; set; }
        public bool IsRequiredForForm { get; set; }
        public bool IsContentElement { get; set; }
        public bool IsControlElement { get; set; }
        public string ItemType { get; set; }
        public string Orientation { get; set; }
        public List<string> SupportedPatterns { get; set; } = new List<string>();

        // === Position & Size ===
        public Rect BoundingRectangle { get; set; }
        public System.Windows.Point ClickablePoint { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public Rect ClientRect { get; set; }

        // === Web/HTML Properties ===
        public string TagName { get; set; }
        public string HtmlId { get; set; }
        public string HtmlClassName { get; set; }
        public string InnerText { get; set; }
        public string InnerHTML { get; set; }
        public string OuterHTML { get; set; }
        public string Href { get; set; }
        public string Src { get; set; }
        public string Alt { get; set; }
        public string Title { get; set; }
        public string Role { get; set; }
        public string AriaLabel { get; set; }
        public string AriaDescribedBy { get; set; }
        public string AriaLabelledBy { get; set; }
        public Dictionary<string, string> HtmlAttributes { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> ComputedStyles { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> AriaAttributes { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> DataAttributes { get; set; } = new Dictionary<string, string>();

        // === Selectors ===
        public string XPath { get; set; }
        public string CssSelector { get; set; }
        public string FullXPath { get; set; }
        public string WindowsPath { get; set; } // Parent > Child notation for desktop
        public string AccessiblePath { get; set; }
        public string TreePath { get; set; } // Complete element tree path
        public string ElementPath { get; set; } // Full hierarchical path to element

        // === Hierarchy ===
        public ElementInfo Parent { get; set; }
        public List<ElementInfo> Children { get; set; } = new List<ElementInfo>();
        public int ChildIndex { get; set; }
        public int TreeLevel { get; set; }

        // === Table/Grid Properties ===
        public int RowIndex { get; set; } = -1; // Table row index if element is in a table (-1 if not in table)
        public int ColumnIndex { get; set; } = -1; // Table column index if element is in a table (-1 if not in table)
        public int RowCount { get; set; } = -1; // Total row count in table (-1 if not a table)
        public int ColumnCount { get; set; } = -1; // Total column count in table (-1 if not a table)
        public int RowSpan { get; set; } = 1; // Number of rows this cell spans
        public int ColumnSpan { get; set; } = 1; // Number of columns this cell spans
        public bool IsTableCell { get; set; } // True if element is a table cell
        public bool IsTableHeader { get; set; } // True if element is a table header
        public string TableName { get; set; } // Name/ID of the table containing this element
        public List<string> ColumnHeaders { get; set; } = new List<string>(); // Column header names
        public List<string> RowHeaders { get; set; } = new List<string>(); // Row header names

        public string ParentName { get; set; }
        public string ParentId { get; set; }
        public string ParentClassName { get; set; }

        // === Window Information ===
        public string WindowTitle { get; set; }
        public string WindowClassName { get; set; }
        public IntPtr WindowHandle { get; set; }
        public string ApplicationName { get; set; }
        public string ApplicationPath { get; set; }

        // === MSHTML/IHTMLDocument Properties ===
        public string DocumentTitle { get; set; }
        public string DocumentUrl { get; set; }
        public string DocumentDomain { get; set; }
        public string DocumentReadyState { get; set; }

        // === Playwright Specific ===
        public string PlaywrightSelector { get; set; }
        public string PlaywrightTableSelector { get; set; } // Table row selector for Playwright
        public bool IsVisible { get; set; }
        public bool IsHidden { get; set; }
        public bool IsChecked { get; set; }
        public bool IsDisabled { get; set; }
        public bool IsEditable { get; set; }
        public string InputType { get; set; }
        public string InputValue { get; set; }

        // === Legacy/Accessibility Properties ===
        public string LegacyName { get; set; }
        public string LegacyValue { get; set; }
        public string LegacyDescription { get; set; }
        public string LegacyHelp { get; set; }
        public string LegacyKeyboardShortcut { get; set; }
        public int LegacyState { get; set; }
        public int LegacyRole { get; set; }

        // === Additional Metadata ===
        public Dictionary<string, object> CustomProperties { get; set; } = new Dictionary<string, object>();
        public List<string> EventListeners { get; set; } = new List<string>();
        public string SourceCode { get; set; }
        public byte[] Screenshot { get; set; }
        public string ScreenshotPath { get; set; }

        // === Collection Profile ===
        public string CollectionProfile { get; set; } // Quick, Standard, Full, Custom
        public TimeSpan CollectionDuration { get; set; }
        public List<string> CollectionErrors { get; set; } = new List<string>();

        /// <summary>
        /// Returns a formatted string representation of all properties
        /// </summary>
        public string ToDetailedString()
        {
            var result = new System.Text.StringBuilder();
            result.AppendLine($"=== Element Information ===");
            result.AppendLine($"Capture Time: {CaptureTime}");
            result.AppendLine($"Detection Method: {DetectionMethod}");
            result.AppendLine($"Collection Profile: {CollectionProfile}");

            if (!string.IsNullOrEmpty(ElementType))
                result.AppendLine($"Element Type: {ElementType}");
            if (!string.IsNullOrEmpty(Name))
                result.AppendLine($"Name: {Name}");
            if (!string.IsNullOrEmpty(AutomationId))
                result.AppendLine($"AutomationId: {AutomationId}");
            // ... Add more properties as needed

            return result.ToString();
        }
    }
}