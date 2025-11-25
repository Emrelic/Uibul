using System;
using System.Collections.Generic;
using System.Windows;

namespace UIElementInspector.Core.Models
{
    /// <summary>
    /// Universal element information container that holds data from all detection technologies
    /// Supports: UI Automation, WebView2/CDP, MSHTML, Win32 API, Playwright
    /// </summary>
    public class ElementInfo
    {
        public Guid Id { get; set; } = Guid.NewGuid();
        public DateTime CaptureTime { get; set; } = DateTime.Now;
        public string DetectionMethod { get; set; } // UI Automation, WebView2, MSHTML, Win32, Playwright, Combined

        #region === 1. BASIC/UNIVERSAL PROPERTIES ===
        // Properties that exist across all technologies
        public string ElementType { get; set; }
        public string Name { get; set; }
        public string ClassName { get; set; }
        public string Value { get; set; }
        public string Description { get; set; }
        public string Role { get; set; } // ARIA role or accessibility role
        public string Tag { get; set; } // HTML tag or control type
        #endregion

        #region === 2. UI AUTOMATION (UIA) PROPERTIES ===
        // 2.1 Basic UIA Properties (AutomationElement.Current)
        public string AutomationId { get; set; }
        public string ControlType { get; set; }
        public string LocalizedControlType { get; set; }
        public string FrameworkId { get; set; } // Win32 / WPF / WinForms / XAML
        public int ProcessId { get; set; }
        public string RuntimeId { get; set; }
        public IntPtr NativeWindowHandle { get; set; }
        public string ItemStatus { get; set; }
        public string ItemType { get; set; }
        public string HelpText { get; set; }
        public string AcceleratorKey { get; set; }
        public string AccessKey { get; set; }
        public string Orientation { get; set; } // Horizontal / Vertical
        public string Culture { get; set; } // Language info
        public string ProviderDescription { get; set; } // UI provider description
        public string LabeledBy { get; set; } // Associated label element

        // 2.2 UIA State Properties
        public bool IsEnabled { get; set; } = true;
        public bool IsOffscreen { get; set; }
        public bool HasKeyboardFocus { get; set; }
        public bool IsKeyboardFocusable { get; set; }
        public bool IsPassword { get; set; }
        public bool IsRequiredForForm { get; set; }
        public bool IsContentElement { get; set; } = true;
        public bool IsControlElement { get; set; } = true;

        // 2.3 UIA Control Patterns & Pattern Properties
        public List<string> SupportedPatterns { get; set; } = new List<string>();

        // ValuePattern
        public bool? ValuePattern_IsReadOnly { get; set; }
        public string ValuePattern_Value { get; set; }

        // RangeValuePattern
        public double? RangeValue_Minimum { get; set; }
        public double? RangeValue_Maximum { get; set; }
        public double? RangeValue_Value { get; set; }
        public double? RangeValue_SmallChange { get; set; }
        public double? RangeValue_LargeChange { get; set; }
        public bool? RangeValue_IsReadOnly { get; set; }

        // TextPattern
        public string TextPattern_DocumentRange { get; set; }
        public string TextPattern_SupportedTextSelection { get; set; }
        public string TextPattern_CaretRange { get; set; }

        // SelectionPattern
        public bool? Selection_CanSelectMultiple { get; set; }
        public bool? Selection_IsSelectionRequired { get; set; }
        public List<string> Selection_SelectedItems { get; set; } = new List<string>();

        // SelectionItemPattern
        public bool? SelectionItem_IsSelected { get; set; }
        public string SelectionItem_Container { get; set; }

        // ScrollPattern
        public double? Scroll_HorizontalPercent { get; set; }
        public double? Scroll_VerticalPercent { get; set; }
        public double? Scroll_HorizontalViewSize { get; set; }
        public double? Scroll_VerticalViewSize { get; set; }
        public bool? Scroll_HorizontallyScrollable { get; set; }
        public bool? Scroll_VerticallyScrollable { get; set; }

        // ExpandCollapsePattern
        public string ExpandCollapse_State { get; set; } // Collapsed, Expanded, PartiallyExpanded, LeafNode

        // TogglePattern
        public string Toggle_State { get; set; } // On, Off, Indeterminate

        // TransformPattern
        public bool? Transform_CanMove { get; set; }
        public bool? Transform_CanResize { get; set; }
        public bool? Transform_CanRotate { get; set; }

        // DockPattern
        public string Dock_Position { get; set; } // Top, Left, Bottom, Right, Fill, None

        // WindowPattern
        public bool? Window_CanMaximize { get; set; }
        public bool? Window_CanMinimize { get; set; }
        public bool? Window_IsModal { get; set; }
        public bool? Window_IsTopmost { get; set; }
        public string Window_InteractionState { get; set; }
        public string Window_VisualState { get; set; }

        // 2.4 LegacyIAccessiblePattern Properties
        public string LegacyName { get; set; }
        public string LegacyValue { get; set; }
        public string LegacyDescription { get; set; }
        public string LegacyHelp { get; set; }
        public string LegacyKeyboardShortcut { get; set; }
        public string LegacyDefaultAction { get; set; }
        public int LegacyState { get; set; }
        public int LegacyRole { get; set; }
        public int LegacyChildCount { get; set; }
        #endregion

        #region === 3. WEBVIEW2 / CDP (Chrome DevTools Protocol) PROPERTIES ===
        // 3.1 HTML DOM Attributes
        public string TagName { get; set; }
        public string HtmlId { get; set; }
        public string HtmlClassName { get; set; }
        public List<string> ClassList { get; set; } = new List<string>();
        public string HtmlName { get; set; }
        public string InnerText { get; set; }
        public string InnerHTML { get; set; }
        public string OuterHTML { get; set; }
        public string Href { get; set; }
        public string Src { get; set; }
        public string Alt { get; set; }
        public string Title { get; set; }
        public string InputType { get; set; }
        public string InputValue { get; set; }
        public string Placeholder { get; set; }
        public int? TabIndex { get; set; }
        public Dictionary<string, string> HtmlAttributes { get; set; } = new Dictionary<string, string>();
        public Dictionary<string, string> DataAttributes { get; set; } = new Dictionary<string, string>(); // data-* attributes

        // 3.2 ARIA Accessibility Attributes
        public string AriaLabel { get; set; }
        public string AriaDescribedBy { get; set; }
        public string AriaLabelledBy { get; set; }
        public string AriaRole { get; set; }
        public bool? AriaDisabled { get; set; }
        public bool? AriaHidden { get; set; }
        public bool? AriaExpanded { get; set; }
        public bool? AriaSelected { get; set; }
        public bool? AriaChecked { get; set; }
        public bool? AriaRequired { get; set; }
        public string AriaHasPopup { get; set; }
        public int? AriaLevel { get; set; }
        public string AriaValueText { get; set; }
        public Dictionary<string, string> AriaAttributes { get; set; } = new Dictionary<string, string>();

        // 3.3 Layout / Box Model Properties
        public double ClientWidth { get; set; }
        public double ClientHeight { get; set; }
        public double OffsetWidth { get; set; }
        public double OffsetHeight { get; set; }
        public double ScrollWidth { get; set; }
        public double ScrollHeight { get; set; }
        public double ScrollTop { get; set; }
        public double ScrollLeft { get; set; }
        public string BoxModel_Margin { get; set; }
        public string BoxModel_Padding { get; set; }
        public string BoxModel_Border { get; set; }

        // 3.4 CSS Computed Styles
        public Dictionary<string, string> ComputedStyles { get; set; } = new Dictionary<string, string>();
        public string Style_Display { get; set; }
        public string Style_Position { get; set; }
        public string Style_Visibility { get; set; }
        public string Style_Opacity { get; set; }
        public string Style_Color { get; set; }
        public string Style_BackgroundColor { get; set; }
        public string Style_FontSize { get; set; }
        public string Style_FontWeight { get; set; }
        public string Style_ZIndex { get; set; }
        public string Style_PointerEvents { get; set; }
        public string Style_Overflow { get; set; }
        public string Style_Transform { get; set; }

        // 3.5 CDP Accessibility Tree (AXNode) Properties
        public string AX_Name { get; set; }
        public string AX_Role { get; set; }
        public string AX_Description { get; set; }
        public string AX_Value { get; set; }
        public string AX_ValueText { get; set; }
        public bool? AX_Focused { get; set; }
        public bool? AX_Disabled { get; set; }
        public bool? AX_Multiline { get; set; }
        public bool? AX_Readonly { get; set; }
        public bool? AX_Selected { get; set; }
        public bool? AX_Checked { get; set; }
        public bool? AX_Required { get; set; }
        public bool? AX_Expanded { get; set; }
        public string AX_HasPopup { get; set; }
        public int? AX_Level { get; set; }

        // 3.6 CDP Performance/Debug Info
        public string CDP_ConsoleMessages { get; set; }
        public double? CDP_DOMContentLoaded { get; set; }
        public double? CDP_LoadEventEnd { get; set; }
        public long? CDP_JSHeapUsed { get; set; }
        public long? CDP_JSHeapLimit { get; set; }
        public int? CDP_NetworkRequestsCount { get; set; }
        #endregion

        #region === 4. MSHTML / IHTMLDocument (Trident/IE) PROPERTIES ===
        // 4.1 IHTMLElement Properties
        public string MSHTML_ParentElementTag { get; set; }
        public int? MSHTML_ChildrenCount { get; set; }
        public string MSHTML_OuterText { get; set; }
        public string MSHTML_Language { get; set; }
        public string MSHTML_Lang { get; set; }
        public int? MSHTML_SourceIndex { get; set; } // Position in document.all collection
        public string MSHTML_ScopeName { get; set; } // Namespace prefix
        public bool? MSHTML_CanHaveChildren { get; set; }
        public bool? MSHTML_CanHaveHTML { get; set; }
        public bool? MSHTML_IsContentEditable { get; set; }
        public bool? MSHTML_HideFocus { get; set; }
        public string MSHTML_ContentEditable { get; set; }
        public int? MSHTML_TabIndex { get; set; }
        public string MSHTML_AccessKey { get; set; }

        // 4.2 IHTMLElement2 Properties
        public double? MSHTML_ClientLeft { get; set; }
        public double? MSHTML_ClientTop { get; set; }
        public string MSHTML_CurrentStyle { get; set; }
        public string MSHTML_RuntimeStyle { get; set; }
        public string MSHTML_ReadyState { get; set; }
        public string MSHTML_Dir { get; set; } // Text direction (ltr/rtl)
        public int? MSHTML_ScrollLeft { get; set; }
        public int? MSHTML_ScrollTop { get; set; }
        public int? MSHTML_ScrollWidth { get; set; }
        public int? MSHTML_ScrollHeight { get; set; }

        // 4.3 IHTMLInputElement Properties
        public string MSHTML_DefaultValue { get; set; }
        public bool? MSHTML_DefaultChecked { get; set; }
        public string MSHTML_Form_Id { get; set; } // Parent form ID
        public string MSHTML_Form_Name { get; set; } // Parent form Name
        public string MSHTML_Form_Action { get; set; } // Parent form action URL
        public string MSHTML_Form_Method { get; set; } // Parent form method (GET/POST)
        public int? MSHTML_MaxLength { get; set; }
        public int? MSHTML_Size { get; set; }
        public bool? MSHTML_ReadOnly { get; set; }
        public string MSHTML_Accept { get; set; } // File input accept types
        public string MSHTML_Align { get; set; }
        public string MSHTML_UseMap { get; set; }
        public bool? MSHTML_IndeterminateState { get; set; }

        // 4.4 IHTMLSelectElement Properties
        public int? MSHTML_SelectedIndex { get; set; }
        public int? MSHTML_OptionsLength { get; set; }
        public List<string> MSHTML_Options { get; set; } = new List<string>();
        public List<string> MSHTML_OptionValues { get; set; } = new List<string>(); // Option value attributes
        public string MSHTML_SelectedValue { get; set; } // Currently selected value
        public string MSHTML_SelectedText { get; set; } // Currently selected text
        public bool? MSHTML_Multiple { get; set; } // Allow multiple selection

        // 4.5 IHTMLTextAreaElement Properties
        public int? MSHTML_Cols { get; set; }
        public int? MSHTML_Rows { get; set; }
        public string MSHTML_Wrap { get; set; } // Text wrap mode

        // 4.6 IHTMLButtonElement Properties
        public string MSHTML_ButtonType { get; set; } // button, submit, reset
        public string MSHTML_FormAction { get; set; } // Button-specific form action
        public string MSHTML_FormMethod { get; set; } // Button-specific form method

        // 4.7 IHTMLAnchorElement Properties
        public string MSHTML_Target { get; set; } // _blank, _self, _parent, _top
        public string MSHTML_Protocol { get; set; } // http:, https:, javascript:
        public string MSHTML_Host { get; set; }
        public string MSHTML_Hostname { get; set; }
        public string MSHTML_Port { get; set; }
        public string MSHTML_Pathname { get; set; }
        public string MSHTML_Search { get; set; } // Query string
        public string MSHTML_Hash { get; set; } // Fragment identifier
        public string MSHTML_Rel { get; set; } // Relationship

        // 4.8 IHTMLImageElement Properties
        public bool? MSHTML_IsMap { get; set; }
        public int? MSHTML_NaturalWidth { get; set; }
        public int? MSHTML_NaturalHeight { get; set; }
        public bool? MSHTML_Complete { get; set; } // Image fully loaded
        public string MSHTML_LongDesc { get; set; }
        public int? MSHTML_Vspace { get; set; }
        public int? MSHTML_Hspace { get; set; }
        public string MSHTML_LowSrc { get; set; }

        // 4.9 IHTMLTableElement Properties
        public string MSHTML_Table_Caption { get; set; }
        public string MSHTML_Table_Summary { get; set; }
        public string MSHTML_Table_Border { get; set; }
        public string MSHTML_Table_CellPadding { get; set; }
        public string MSHTML_Table_CellSpacing { get; set; }
        public string MSHTML_Table_Width { get; set; }
        public string MSHTML_Table_BgColor { get; set; }
        public int? MSHTML_Table_RowsCount { get; set; }
        public int? MSHTML_Table_TBodiesCount { get; set; }

        // 4.10 IHTMLTableRowElement Properties
        public int? MSHTML_Row_RowIndex { get; set; }
        public int? MSHTML_Row_SectionRowIndex { get; set; }
        public int? MSHTML_Row_CellsCount { get; set; }
        public string MSHTML_Row_BgColor { get; set; }
        public string MSHTML_Row_VAlign { get; set; }
        public string MSHTML_Row_Align { get; set; }

        // 4.11 IHTMLTableCellElement Properties
        public int? MSHTML_Cell_CellIndex { get; set; }
        public string MSHTML_Cell_Abbr { get; set; }
        public string MSHTML_Cell_Axis { get; set; }
        public string MSHTML_Cell_Headers { get; set; }
        public string MSHTML_Cell_Scope { get; set; } // row, col, rowgroup, colgroup
        public string MSHTML_Cell_NoWrap { get; set; }
        public string MSHTML_Cell_BgColor { get; set; }
        public string MSHTML_Cell_VAlign { get; set; }
        public string MSHTML_Cell_Align { get; set; }

        // 4.12 IHTMLFrameElement / IHTMLIFrameElement Properties
        public string MSHTML_Frame_Src { get; set; }
        public string MSHTML_Frame_Name { get; set; }
        public string MSHTML_Frame_Scrolling { get; set; }
        public string MSHTML_Frame_FrameBorder { get; set; }
        public string MSHTML_Frame_MarginWidth { get; set; }
        public string MSHTML_Frame_MarginHeight { get; set; }
        public bool? MSHTML_Frame_NoResize { get; set; }

        // 4.13 IHTMLDocument Properties
        public string DocumentTitle { get; set; }
        public string DocumentUrl { get; set; }
        public string DocumentDomain { get; set; }
        public string DocumentReadyState { get; set; }
        public string DocumentCookie { get; set; }
        public int? DocumentFramesCount { get; set; }
        public int? DocumentScriptsCount { get; set; }
        public int? DocumentLinksCount { get; set; }
        public int? DocumentImagesCount { get; set; }
        public int? DocumentFormsCount { get; set; }
        public string DocumentActiveElement { get; set; }
        public string DocumentCharset { get; set; }
        public string DocumentLastModified { get; set; }
        public string DocumentReferrer { get; set; }
        public string DocumentCompatMode { get; set; } // BackCompat or CSS1Compat
        public string DocumentDesignMode { get; set; } // on/off
        public string DocumentDocType { get; set; }
        public string DocumentDir { get; set; } // Document text direction

        // 4.14 IHTMLDocument2/3/4/5 Additional Properties
        public string MSHTML_Doc_Protocol { get; set; }
        public string MSHTML_Doc_NameProp { get; set; }
        public string MSHTML_Doc_FileCreatedDate { get; set; }
        public string MSHTML_Doc_FileModifiedDate { get; set; }
        public string MSHTML_Doc_FileSize { get; set; }
        public string MSHTML_Doc_MimeType { get; set; }
        public string MSHTML_Doc_Security { get; set; }
        public int? MSHTML_Doc_Anchors_Count { get; set; }
        public int? MSHTML_Doc_Applets_Count { get; set; }
        public int? MSHTML_Doc_Embeds_Count { get; set; }
        public int? MSHTML_Doc_All_Count { get; set; } // Total element count
        #endregion

        #region === 5. WIN32 API / SPY++ PROPERTIES ===
        // 5.1 Window Handle Information
        public IntPtr Win32_HWND { get; set; }
        public IntPtr Win32_ParentHWND { get; set; }
        public List<IntPtr> Win32_ChildHWNDs { get; set; } = new List<IntPtr>();
        public string Win32_ClassName { get; set; } // WNDCLASS name (WC_BUTTON, EDIT, STATIC, etc.)
        public string Win32_WindowText { get; set; } // Caption / Text

        // 5.2 Process/Thread Information
        public int Win32_ThreadId { get; set; }
        public IntPtr Win32_InstanceHandle { get; set; }
        public IntPtr Win32_MenuHandle { get; set; }
        public IntPtr Win32_WndProc { get; set; }

        // 5.3 Window Styles (WS_* flags)
        public uint Win32_WindowStyles { get; set; }
        public string Win32_WindowStyles_Parsed { get; set; }

        // 5.4 Extended Window Styles (WS_EX_* flags)
        public uint Win32_ExtendedStyles { get; set; }
        public string Win32_ExtendedStyles_Parsed { get; set; }

        // 5.5 Window Rectangles
        public Rect Win32_WindowRect { get; set; }
        public Rect Win32_ClientRect { get; set; }

        // 5.6 Control-Specific Information
        public string Win32_ControlId { get; set; }
        public bool? Win32_IsUnicode { get; set; }
        public bool? Win32_IsVisible { get; set; }
        public bool? Win32_IsEnabled { get; set; }
        public bool? Win32_IsMaximized { get; set; }
        public bool? Win32_IsMinimized { get; set; }
        #endregion

        #region === 6. POSITION & SIZE (Universal) ===
        public Rect BoundingRectangle { get; set; }
        public System.Windows.Point ClickablePoint { get; set; }
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public Rect ClientRect { get; set; }
        #endregion

        #region === 7. SELECTORS ===
        public string XPath { get; set; }
        public string CssSelector { get; set; }
        public string FullXPath { get; set; }
        public string WindowsPath { get; set; } // Parent > Child notation for desktop
        public string AccessiblePath { get; set; }
        public string TreePath { get; set; } // Complete element tree path
        public string ElementPath { get; set; } // Full hierarchical path to element
        public string PlaywrightSelector { get; set; }
        public string PlaywrightTableSelector { get; set; }
        #endregion

        #region === 8. HIERARCHY ===
        public ElementInfo Parent { get; set; }
        public List<ElementInfo> Children { get; set; } = new List<ElementInfo>();
        public ElementInfo NextSibling { get; set; }
        public ElementInfo PreviousSibling { get; set; }
        public int ChildIndex { get; set; }
        public int TreeLevel { get; set; }
        public string ParentName { get; set; }
        public string ParentId { get; set; }
        public string ParentClassName { get; set; }
        public string OwnerWindow { get; set; }
        public string ControlContainer { get; set; }
        #endregion

        #region === 9. TABLE/GRID PROPERTIES ===
        public int RowIndex { get; set; } = -1;
        public int ColumnIndex { get; set; } = -1;
        public int RowCount { get; set; } = -1;
        public int ColumnCount { get; set; } = -1;
        public int RowSpan { get; set; } = 1;
        public int ColumnSpan { get; set; } = 1;
        public bool IsTableCell { get; set; }
        public bool IsTableHeader { get; set; }
        public string TableName { get; set; }
        public List<string> ColumnHeaders { get; set; } = new List<string>();
        public List<string> RowHeaders { get; set; } = new List<string>();
        public string Table_RowOrColumnMajor { get; set; }
        #endregion

        #region === 10. WINDOW/APPLICATION INFO ===
        public string WindowTitle { get; set; }
        public string WindowClassName { get; set; }
        public IntPtr WindowHandle { get; set; }
        public string ApplicationName { get; set; }
        public string ApplicationPath { get; set; }
        #endregion

        #region === 11. STATE PROPERTIES (Universal) ===
        public bool IsVisible { get; set; } = true;
        public bool IsHidden { get; set; }
        public bool IsChecked { get; set; }
        public bool IsDisabled { get; set; }
        public bool IsEditable { get; set; }
        public bool IsSelected { get; set; }
        public bool IsFocused { get; set; }
        public bool IsExpanded { get; set; }
        #endregion

        #region === 12. ADDITIONAL METADATA ===
        public Dictionary<string, object> CustomProperties { get; set; } = new Dictionary<string, object>();
        public List<string> EventListeners { get; set; } = new List<string>();
        public string SourceCode { get; set; }
        public byte[] Screenshot { get; set; }
        public string ScreenshotPath { get; set; }
        #endregion

        #region === 13. COLLECTION INFO ===
        public string CollectionProfile { get; set; } // Quick, Standard, Full, Custom
        public TimeSpan CollectionDuration { get; set; }
        public List<string> CollectionErrors { get; set; } = new List<string>();
        public List<string> TechnologiesUsed { get; set; } = new List<string>();
        #endregion

        /// <summary>
        /// Returns a formatted string representation of all properties organized by technology
        /// </summary>
        public string ToDetailedString()
        {
            var result = new System.Text.StringBuilder();
            result.AppendLine("╔════════════════════════════════════════════════════════════════╗");
            result.AppendLine("║              UNIVERSAL UI ELEMENT INSPECTOR                     ║");
            result.AppendLine("╚════════════════════════════════════════════════════════════════╝");
            result.AppendLine();
            result.AppendLine($"📅 Capture Time: {CaptureTime}");
            result.AppendLine($"🔍 Detection Method: {DetectionMethod}");
            result.AppendLine($"📊 Collection Profile: {CollectionProfile}");
            result.AppendLine($"⏱️ Collection Duration: {CollectionDuration.TotalMilliseconds}ms");
            result.AppendLine();

            // Basic Properties
            result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
            result.AppendLine("│ 1. BASIC / UNIVERSAL PROPERTIES                                 │");
            result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
            AppendIfNotNull(result, "Element Type", ElementType);
            AppendIfNotNull(result, "Name", Name);
            AppendIfNotNull(result, "Class Name", ClassName);
            AppendIfNotNull(result, "Value", Value);
            AppendIfNotNull(result, "Description", Description);
            AppendIfNotNull(result, "Role", Role);
            result.AppendLine();

            // UI Automation Properties
            result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
            result.AppendLine("│ 2. UI AUTOMATION (UIA) PROPERTIES                               │");
            result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
            AppendIfNotNull(result, "AutomationId", AutomationId);
            AppendIfNotNull(result, "Control Type", ControlType);
            AppendIfNotNull(result, "Localized Control Type", LocalizedControlType);
            AppendIfNotNull(result, "Framework Id", FrameworkId);
            result.AppendLine($"  Process Id: {ProcessId}");
            AppendIfNotNull(result, "Runtime Id", RuntimeId);
            result.AppendLine($"  Is Enabled: {IsEnabled}");
            result.AppendLine($"  Is Offscreen: {IsOffscreen}");
            result.AppendLine($"  Has Keyboard Focus: {HasKeyboardFocus}");
            result.AppendLine($"  Is Keyboard Focusable: {IsKeyboardFocusable}");
            AppendIfNotNull(result, "Help Text", HelpText);
            if (SupportedPatterns.Count > 0)
                result.AppendLine($"  Supported Patterns: {string.Join(", ", SupportedPatterns)}");
            result.AppendLine();

            // LegacyIAccessible
            if (LegacyRole > 0 || !string.IsNullOrEmpty(LegacyName))
            {
                result.AppendLine("  ── LegacyIAccessible ──");
                AppendIfNotNull(result, "Legacy Name", LegacyName);
                AppendIfNotNull(result, "Legacy Value", LegacyValue);
                AppendIfNotNull(result, "Legacy Description", LegacyDescription);
                AppendIfNotNull(result, "Legacy Default Action", LegacyDefaultAction);
                AppendIfNotNull(result, "Legacy Keyboard Shortcut", LegacyKeyboardShortcut);
                result.AppendLine($"  Legacy Role: {LegacyRole}");
                result.AppendLine($"  Legacy State: {LegacyState}");
                result.AppendLine($"  Legacy Child Count: {LegacyChildCount}");
            }
            result.AppendLine();

            // Web/CDP Properties
            if (!string.IsNullOrEmpty(TagName))
            {
                result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
                result.AppendLine("│ 3. WEBVIEW2 / CDP PROPERTIES                                    │");
                result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
                AppendIfNotNull(result, "Tag Name", TagName);
                AppendIfNotNull(result, "HTML Id", HtmlId);
                AppendIfNotNull(result, "HTML Class", HtmlClassName);
                AppendIfNotNull(result, "HTML Name", HtmlName);
                AppendIfNotNull(result, "Inner Text", InnerText);
                AppendIfNotNull(result, "Href", Href);
                AppendIfNotNull(result, "Src", Src);
                AppendIfNotNull(result, "Input Type", InputType);
                AppendIfNotNull(result, "Placeholder", Placeholder);

                if (AriaAttributes.Count > 0 || !string.IsNullOrEmpty(AriaLabel))
                {
                    result.AppendLine("  ── ARIA Attributes ──");
                    AppendIfNotNull(result, "ARIA Label", AriaLabel);
                    AppendIfNotNull(result, "ARIA Role", AriaRole);
                    AppendIfNotNull(result, "ARIA Described By", AriaDescribedBy);
                    foreach (var aria in AriaAttributes)
                        result.AppendLine($"  {aria.Key}: {aria.Value}");
                }

                if (ComputedStyles.Count > 0)
                {
                    result.AppendLine("  ── Computed Styles ──");
                    AppendIfNotNull(result, "Display", Style_Display);
                    AppendIfNotNull(result, "Position", Style_Position);
                    AppendIfNotNull(result, "Visibility", Style_Visibility);
                    AppendIfNotNull(result, "Color", Style_Color);
                    AppendIfNotNull(result, "Background", Style_BackgroundColor);
                }
                result.AppendLine();
            }

            // MSHTML Properties
            if (!string.IsNullOrEmpty(DocumentUrl))
            {
                result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
                result.AppendLine("│ 4. MSHTML / IHTMLDocument PROPERTIES                            │");
                result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
                AppendIfNotNull(result, "Document Title", DocumentTitle);
                AppendIfNotNull(result, "Document URL", DocumentUrl);
                AppendIfNotNull(result, "Document Domain", DocumentDomain);
                AppendIfNotNull(result, "Ready State", DocumentReadyState);
                result.AppendLine();
            }

            // Win32 Properties
            if (Win32_HWND != IntPtr.Zero)
            {
                result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
                result.AppendLine("│ 5. WIN32 API / SPY++ PROPERTIES                                 │");
                result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
                result.AppendLine($"  HWND: 0x{Win32_HWND.ToInt64():X8}");
                result.AppendLine($"  Parent HWND: 0x{Win32_ParentHWND.ToInt64():X8}");
                AppendIfNotNull(result, "Win32 Class Name", Win32_ClassName);
                AppendIfNotNull(result, "Win32 Window Text", Win32_WindowText);
                result.AppendLine($"  Thread Id: {Win32_ThreadId}");
                AppendIfNotNull(result, "Window Styles", Win32_WindowStyles_Parsed);
                AppendIfNotNull(result, "Extended Styles", Win32_ExtendedStyles_Parsed);
                result.AppendLine();
            }

            // Position & Size
            result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
            result.AppendLine("│ 6. POSITION & SIZE                                              │");
            result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
            result.AppendLine($"  Bounding Rectangle: {BoundingRectangle}");
            result.AppendLine($"  X: {X}, Y: {Y}");
            result.AppendLine($"  Width: {Width}, Height: {Height}");
            result.AppendLine($"  Clickable Point: {ClickablePoint}");
            result.AppendLine();

            // Selectors
            result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
            result.AppendLine("│ 7. SELECTORS                                                    │");
            result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
            AppendIfNotNull(result, "XPath", XPath);
            AppendIfNotNull(result, "CSS Selector", CssSelector);
            AppendIfNotNull(result, "Windows Path", WindowsPath);
            AppendIfNotNull(result, "Tree Path", TreePath);
            AppendIfNotNull(result, "Playwright Selector", PlaywrightSelector);
            result.AppendLine();

            // Hierarchy
            result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
            result.AppendLine("│ 8. HIERARCHY                                                    │");
            result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
            AppendIfNotNull(result, "Parent Name", ParentName);
            AppendIfNotNull(result, "Parent Id", ParentId);
            AppendIfNotNull(result, "Parent Class", ParentClassName);
            result.AppendLine($"  Children Count: {Children?.Count ?? 0}");
            result.AppendLine($"  Tree Level: {TreeLevel}");
            result.AppendLine();

            // Table/Grid
            if (IsTableCell || RowIndex >= 0)
            {
                result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
                result.AppendLine("│ 9. TABLE/GRID PROPERTIES                                        │");
                result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
                result.AppendLine($"  Is Table Cell: {IsTableCell}");
                result.AppendLine($"  Is Table Header: {IsTableHeader}");
                result.AppendLine($"  Row Index: {RowIndex}");
                result.AppendLine($"  Column Index: {ColumnIndex}");
                result.AppendLine($"  Row Count: {RowCount}");
                result.AppendLine($"  Column Count: {ColumnCount}");
                AppendIfNotNull(result, "Table Name", TableName);
                if (ColumnHeaders.Count > 0)
                    result.AppendLine($"  Column Headers: {string.Join(", ", ColumnHeaders)}");
                result.AppendLine();
            }

            // Custom Properties
            if (CustomProperties.Count > 0)
            {
                result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
                result.AppendLine("│ CUSTOM PROPERTIES                                               │");
                result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
                foreach (var prop in CustomProperties)
                {
                    result.AppendLine($"  {prop.Key}: {prop.Value}");
                }
                result.AppendLine();
            }

            // Collection Errors
            if (CollectionErrors.Count > 0)
            {
                result.AppendLine("┌─────────────────────────────────────────────────────────────────┐");
                result.AppendLine("│ COLLECTION NOTES/ERRORS                                         │");
                result.AppendLine("└─────────────────────────────────────────────────────────────────┘");
                foreach (var error in CollectionErrors)
                {
                    result.AppendLine($"  ⚠ {error}");
                }
            }

            return result.ToString();
        }

        private void AppendIfNotNull(System.Text.StringBuilder sb, string name, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                sb.AppendLine($"  {name}: {value}");
            }
        }
    }
}