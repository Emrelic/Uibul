using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Detectors
{
    /// <summary>
    /// Detects elements using UI Automation API for native Windows applications
    /// </summary>
    public class UIAutomationDetector : IElementDetector
    {
        public string Name => "UI Automation";

        public bool CanDetect(System.Windows.Point screenPoint)
        {
            try
            {
                var element = AutomationElement.FromPoint(new System.Windows.Point(screenPoint.X, screenPoint.Y));
                return element != null;
            }
            catch
            {
                return false;
            }
        }

        public async Task<ElementInfo> GetElementAtPoint(System.Windows.Point screenPoint, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var element = AutomationElement.FromPoint(new System.Windows.Point(screenPoint.X, screenPoint.Y));
                    if (element == null) return null;

                    return ExtractElementInfo(element, profile);
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"UIAutomation detection error: {ex.Message}");
                    return null;
                }
            });
        }

        public async Task<List<ElementInfo>> GetAllElements(IntPtr windowHandle, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                var elements = new List<ElementInfo>();
                try
                {
                    AutomationElement rootElement;
                    if (windowHandle == IntPtr.Zero)
                    {
                        rootElement = AutomationElement.RootElement;
                    }
                    else
                    {
                        rootElement = AutomationElement.FromHandle(windowHandle);
                    }

                    if (rootElement != null)
                    {
                        CollectAllElements(rootElement, elements, profile, 0);
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"UIAutomation GetAllElements error: {ex.Message}");
                }
                return elements;
            });
        }

        public async Task<List<ElementInfo>> GetElementsInRegion(Rect region, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                var elements = new List<ElementInfo>();
                try
                {
                    var condition = new System.Windows.Automation.PropertyCondition(AutomationElement.IsOffscreenProperty, false);
                    var allElements = AutomationElement.RootElement.FindAll(TreeScope.Descendants, condition);

                    foreach (AutomationElement element in allElements)
                    {
                        var bounds = element.Current.BoundingRectangle;
                        if (!bounds.IsEmpty && region.Contains(bounds))
                        {
                            var info = ExtractElementInfo(element, profile);
                            if (info != null)
                                elements.Add(info);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"UIAutomation GetElementsInRegion error: {ex.Message}");
                }
                return elements;
            });
        }

        public async Task<ElementInfo> GetElementTree(ElementInfo rootElement, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                try
                {
                    if (rootElement?.NativeWindowHandle == IntPtr.Zero)
                        return rootElement;

                    var automationElement = AutomationElement.FromHandle(rootElement.NativeWindowHandle);
                    if (automationElement != null)
                    {
                        var tree = BuildElementTree(automationElement, profile, 0);
                        return tree;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"UIAutomation GetElementTree error: {ex.Message}");
                }
                return rootElement;
            });
        }

        public async Task<ElementInfo> RefreshElement(ElementInfo element, CollectionProfile profile)
        {
            return await Task.Run(() =>
            {
                try
                {
                    if (element?.RuntimeId != null && !string.IsNullOrEmpty(element.RuntimeId))
                    {
                        var runtimeIdParts = element.RuntimeId.Split('.').Select(int.Parse).ToArray();
                        var automationElement = AutomationElement.RootElement.FindFirst(
                            TreeScope.Descendants,
                            new System.Windows.Automation.PropertyCondition(AutomationElement.RuntimeIdProperty, runtimeIdParts)
                        );

                        if (automationElement != null)
                        {
                            return ExtractElementInfo(automationElement, profile);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"UIAutomation RefreshElement error: {ex.Message}");
                }
                return element;
            });
        }

        private ElementInfo ExtractElementInfo(AutomationElement element, CollectionProfile profile)
        {
            if (element == null) return null;

            var info = new ElementInfo
            {
                DetectionMethod = Name,
                CollectionProfile = profile.ToString(),
                CaptureTime = DateTime.Now
            };

            var stopwatch = Stopwatch.StartNew();

            try
            {
                // Always collect basic properties
                ExtractBasicProperties(element, info);

                // Collect additional properties based on profile
                switch (profile)
                {
                    case CollectionProfile.Quick:
                        // Only basic properties (already collected)
                        break;

                    case CollectionProfile.Standard:
                        ExtractStandardProperties(element, info);
                        break;

                    case CollectionProfile.Full:
                        ExtractBasicProperties(element, info);
                        ExtractStandardProperties(element, info);
                        ExtractAdvancedProperties(element, info);
                        ExtractPatterns(element, info);
                        ExtractLegacyProperties(element, info);
                        break;

                    case CollectionProfile.Custom:
                        // TODO: Implement custom settings
                        ExtractStandardProperties(element, info);
                        break;
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Extraction error: {ex.Message}");
            }

            stopwatch.Stop();
            info.CollectionDuration = stopwatch.Elapsed;

            return info;
        }

        private void ExtractBasicProperties(AutomationElement element, ElementInfo info)
        {
            try
            {
                var current = element.Current;

                info.Name = current.Name;
                info.ClassName = current.ClassName;
                info.AutomationId = current.AutomationId;
                info.ControlType = current.ControlType?.ProgrammaticName;
                info.LocalizedControlType = current.LocalizedControlType;

                // For web elements, AutomationId might be provided through ValuePattern or other methods
                // Try to get it using GetCurrentPropertyValue as a fallback
                if (string.IsNullOrEmpty(info.AutomationId))
                {
                    try
                    {
                        var automationIdProperty = element.GetCurrentPropertyValue(AutomationElement.AutomationIdProperty) as string;
                        if (!string.IsNullOrEmpty(automationIdProperty))
                        {
                            info.AutomationId = automationIdProperty;
                            Debug.WriteLine($"UIAutomation - Retrieved AutomationId via Property: '{info.AutomationId}'");
                        }
                    }
                    catch { }
                }

                Debug.WriteLine($"UIAutomation Extract - Name: '{info.Name}', ClassName: '{info.ClassName}', AutomationId: '{info.AutomationId}', ControlType: '{info.ControlType}'");

                // Position and size
                if (!current.BoundingRectangle.IsEmpty)
                {
                    info.BoundingRectangle = current.BoundingRectangle;
                    info.X = current.BoundingRectangle.X;
                    info.Y = current.BoundingRectangle.Y;
                    info.Width = current.BoundingRectangle.Width;
                    info.Height = current.BoundingRectangle.Height;
                }

                info.IsEnabled = current.IsEnabled;
                info.IsOffscreen = current.IsOffscreen;
                info.ProcessId = current.ProcessId;
                info.NativeWindowHandle = new IntPtr(current.NativeWindowHandle);
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Basic properties error: {ex.Message}");
            }
        }

        private void ExtractStandardProperties(AutomationElement element, ElementInfo info)
        {
            try
            {
                var current = element.Current;

                info.FrameworkId = current.FrameworkId;
                info.ItemType = current.ItemType;
                info.ItemStatus = current.ItemStatus;
                info.HelpText = current.HelpText;
                info.AcceleratorKey = current.AcceleratorKey;
                info.AccessKey = current.AccessKey;
                info.HasKeyboardFocus = current.HasKeyboardFocus;
                info.IsFocused = current.HasKeyboardFocus;
                info.IsKeyboardFocusable = current.IsKeyboardFocusable;
                info.IsPassword = current.IsPassword;
                info.IsRequiredForForm = current.IsRequiredForForm;
                info.IsContentElement = current.IsContentElement;
                info.IsControlElement = current.IsControlElement;

                // Runtime ID
                try
                {
                    var runtimeId = element.GetRuntimeId();
                    if (runtimeId != null && runtimeId.Length > 0)
                    {
                        info.RuntimeId = string.Join(".", runtimeId.Select(id => id.ToString()));
                    }
                }
                catch { }

                // Clickable point
                System.Windows.Point clickPoint;
                if (element.TryGetClickablePoint(out clickPoint))
                {
                    info.ClickablePoint = clickPoint;
                }

                // Orientation
                var orientation = current.Orientation;
                info.Orientation = orientation.ToString();

                // Culture (Language)
                try
                {
                    var culture = element.GetCurrentPropertyValue(AutomationElement.CultureProperty);
                    if (culture != null)
                        info.Culture = culture.ToString();
                }
                catch { }

                // Provider Description - may not be available in all .NET versions
                // Removed because ProviderDescriptionProperty is not available in standard .NET Framework

                // Labeled By
                try
                {
                    var labeledBy = element.GetCurrentPropertyValue(AutomationElement.LabeledByProperty) as AutomationElement;
                    if (labeledBy != null)
                    {
                        info.LabeledBy = labeledBy.Current.Name ?? labeledBy.Current.AutomationId;
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Standard properties error: {ex.Message}");
            }
        }

        private void ExtractAdvancedProperties(AutomationElement element, ElementInfo info)
        {
            try
            {
                // Get process information
                if (info.ProcessId > 0)
                {
                    try
                    {
                        var process = Process.GetProcessById(info.ProcessId);
                        info.ApplicationName = process.ProcessName;
                        info.ApplicationPath = process.MainModule?.FileName;
                        info.WindowTitle = process.MainWindowTitle;
                    }
                    catch { }
                }

                // Build Windows path (Parent > Child notation)
                BuildWindowsPath(element, info);

                // Get parent information
                try
                {
                    var parent = TreeWalker.RawViewWalker.GetParent(element);
                    if (parent != null)
                    {
                        info.ParentName = parent.Current.Name;
                        info.ParentId = parent.Current.AutomationId;
                        info.ParentClassName = parent.Current.ClassName;
                    }
                }
                catch { }

                // Count children
                try
                {
                    var children = element.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
                    if (children != null)
                    {
                        info.Children = new List<ElementInfo>();
                        // Note: We don't recursively extract children here to avoid performance issues
                        // Children count is enough for standard view
                    }
                }
                catch { }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Advanced properties error: {ex.Message}");
            }
        }

        private void ExtractPatterns(AutomationElement element, ElementInfo info)
        {
            try
            {
                var patterns = element.GetSupportedPatterns();
                foreach (var pattern in patterns)
                {
                    info.SupportedPatterns.Add(pattern.ProgrammaticName);

                    // Extract pattern-specific properties
                    if (pattern == ValuePattern.Pattern)
                    {
                        var valuePattern = element.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                        if (valuePattern != null)
                        {
                            info.Value = valuePattern.Current.Value;
                            info.ValuePattern_Value = valuePattern.Current.Value;
                            info.ValuePattern_IsReadOnly = valuePattern.Current.IsReadOnly;
                            info.CustomProperties["ValuePattern_IsReadOnly"] = valuePattern.Current.IsReadOnly;
                        }
                    }
                    else if (pattern == RangeValuePattern.Pattern)
                    {
                        var rangePattern = element.GetCurrentPattern(RangeValuePattern.Pattern) as RangeValuePattern;
                        if (rangePattern != null)
                        {
                            info.RangeValue_Minimum = rangePattern.Current.Minimum;
                            info.RangeValue_Maximum = rangePattern.Current.Maximum;
                            info.RangeValue_Value = rangePattern.Current.Value;
                            info.RangeValue_SmallChange = rangePattern.Current.SmallChange;
                            info.RangeValue_LargeChange = rangePattern.Current.LargeChange;
                            info.RangeValue_IsReadOnly = rangePattern.Current.IsReadOnly;
                            info.CustomProperties["RangeMin"] = rangePattern.Current.Minimum;
                            info.CustomProperties["RangeMax"] = rangePattern.Current.Maximum;
                            info.CustomProperties["RangeValue"] = rangePattern.Current.Value;
                        }
                    }
                    else if (pattern == SelectionPattern.Pattern)
                    {
                        var selectionPattern = element.GetCurrentPattern(SelectionPattern.Pattern) as SelectionPattern;
                        if (selectionPattern != null)
                        {
                            info.Selection_CanSelectMultiple = selectionPattern.Current.CanSelectMultiple;
                            info.Selection_IsSelectionRequired = selectionPattern.Current.IsSelectionRequired;
                            info.CustomProperties["CanSelectMultiple"] = selectionPattern.Current.CanSelectMultiple;
                            info.CustomProperties["IsSelectionRequired"] = selectionPattern.Current.IsSelectionRequired;

                            // Get selected items
                            try
                            {
                                var selectedItems = selectionPattern.Current.GetSelection();
                                if (selectedItems != null)
                                {
                                    foreach (var item in selectedItems)
                                    {
                                        info.Selection_SelectedItems.Add(item.Current.Name);
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    else if (pattern == SelectionItemPattern.Pattern)
                    {
                        try
                        {
                            var selItemPattern = element.GetCurrentPattern(SelectionItemPattern.Pattern) as SelectionItemPattern;
                            if (selItemPattern != null)
                            {
                                info.SelectionItem_IsSelected = selItemPattern.Current.IsSelected;
                                info.IsSelected = selItemPattern.Current.IsSelected;
                                try
                                {
                                    var container = selItemPattern.Current.SelectionContainer;
                                    if (container != null)
                                        info.SelectionItem_Container = container.Current.Name ?? container.Current.AutomationId;
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                    else if (pattern == TogglePattern.Pattern)
                    {
                        var togglePattern = element.GetCurrentPattern(TogglePattern.Pattern) as TogglePattern;
                        if (togglePattern != null)
                        {
                            info.Toggle_State = togglePattern.Current.ToggleState.ToString();
                            info.IsChecked = togglePattern.Current.ToggleState == ToggleState.On;
                            info.CustomProperties["ToggleState"] = togglePattern.Current.ToggleState.ToString();
                        }
                    }
                    else if (pattern == ExpandCollapsePattern.Pattern)
                    {
                        var expandPattern = element.GetCurrentPattern(ExpandCollapsePattern.Pattern) as ExpandCollapsePattern;
                        if (expandPattern != null)
                        {
                            info.ExpandCollapse_State = expandPattern.Current.ExpandCollapseState.ToString();
                            info.IsExpanded = expandPattern.Current.ExpandCollapseState == ExpandCollapseState.Expanded;
                            info.CustomProperties["ExpandCollapseState"] = expandPattern.Current.ExpandCollapseState.ToString();
                        }
                    }
                    else if (pattern == ScrollPattern.Pattern)
                    {
                        try
                        {
                            var scrollPattern = element.GetCurrentPattern(ScrollPattern.Pattern) as ScrollPattern;
                            if (scrollPattern != null)
                            {
                                info.Scroll_HorizontalPercent = scrollPattern.Current.HorizontalScrollPercent;
                                info.Scroll_VerticalPercent = scrollPattern.Current.VerticalScrollPercent;
                                info.Scroll_HorizontalViewSize = scrollPattern.Current.HorizontalViewSize;
                                info.Scroll_VerticalViewSize = scrollPattern.Current.VerticalViewSize;
                                info.Scroll_HorizontallyScrollable = scrollPattern.Current.HorizontallyScrollable;
                                info.Scroll_VerticallyScrollable = scrollPattern.Current.VerticallyScrollable;
                            }
                        }
                        catch { }
                    }
                    else if (pattern == TransformPattern.Pattern)
                    {
                        try
                        {
                            var transformPattern = element.GetCurrentPattern(TransformPattern.Pattern) as TransformPattern;
                            if (transformPattern != null)
                            {
                                info.Transform_CanMove = transformPattern.Current.CanMove;
                                info.Transform_CanResize = transformPattern.Current.CanResize;
                                info.Transform_CanRotate = transformPattern.Current.CanRotate;
                            }
                        }
                        catch { }
                    }
                    else if (pattern == DockPattern.Pattern)
                    {
                        try
                        {
                            var dockPattern = element.GetCurrentPattern(DockPattern.Pattern) as DockPattern;
                            if (dockPattern != null)
                            {
                                info.Dock_Position = dockPattern.Current.DockPosition.ToString();
                            }
                        }
                        catch { }
                    }
                    else if (pattern == WindowPattern.Pattern)
                    {
                        try
                        {
                            var windowPattern = element.GetCurrentPattern(WindowPattern.Pattern) as WindowPattern;
                            if (windowPattern != null)
                            {
                                info.Window_CanMaximize = windowPattern.Current.CanMaximize;
                                info.Window_CanMinimize = windowPattern.Current.CanMinimize;
                                info.Window_IsModal = windowPattern.Current.IsModal;
                                info.Window_IsTopmost = windowPattern.Current.IsTopmost;
                                info.Window_InteractionState = windowPattern.Current.WindowInteractionState.ToString();
                                info.Window_VisualState = windowPattern.Current.WindowVisualState.ToString();
                            }
                        }
                        catch { }
                    }
                    else if (pattern == TextPattern.Pattern)
                    {
                        try
                        {
                            var textPattern = element.GetCurrentPattern(TextPattern.Pattern) as TextPattern;
                            if (textPattern != null)
                            {
                                info.TextPattern_SupportedTextSelection = textPattern.SupportedTextSelection.ToString();
                                try
                                {
                                    var docRange = textPattern.DocumentRange;
                                    if (docRange != null)
                                    {
                                        var text = docRange.GetText(1000);
                                        info.TextPattern_DocumentRange = text;
                                    }
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                    else if (pattern == InvokePattern.Pattern)
                    {
                        // InvokePattern just has Invoke() method, no properties
                        info.CustomProperties["CanInvoke"] = true;
                    }
                    else if (pattern == TablePattern.Pattern)
                    {
                        var tablePattern = element.GetCurrentPattern(TablePattern.Pattern) as TablePattern;
                        if (tablePattern != null)
                        {
                            info.RowCount = tablePattern.Current.RowCount;
                            info.ColumnCount = tablePattern.Current.ColumnCount;
                            info.CustomProperties["TableRowOrColumnMajor"] = tablePattern.Current.RowOrColumnMajor.ToString();

                            // Get column headers
                            var columnHeaders = tablePattern.Current.GetColumnHeaders();
                            if (columnHeaders != null)
                            {
                                foreach (var header in columnHeaders)
                                {
                                    try
                                    {
                                        string headerText = header.Current.Name;
                                        if (!string.IsNullOrEmpty(headerText))
                                            info.ColumnHeaders.Add(headerText);
                                    }
                                    catch { }
                                }
                            }

                            // Get row headers
                            var rowHeaders = tablePattern.Current.GetRowHeaders();
                            if (rowHeaders != null)
                            {
                                foreach (var header in rowHeaders)
                                {
                                    try
                                    {
                                        string headerText = header.Current.Name;
                                        if (!string.IsNullOrEmpty(headerText))
                                            info.RowHeaders.Add(headerText);
                                    }
                                    catch { }
                                }
                            }
                        }
                    }
                    else if (pattern == GridPattern.Pattern)
                    {
                        var gridPattern = element.GetCurrentPattern(GridPattern.Pattern) as GridPattern;
                        if (gridPattern != null)
                        {
                            // Only set if not already set by TablePattern
                            if (info.RowCount < 0)
                                info.RowCount = gridPattern.Current.RowCount;
                            if (info.ColumnCount < 0)
                                info.ColumnCount = gridPattern.Current.ColumnCount;
                        }
                    }
                    else if (pattern == TableItemPattern.Pattern)
                    {
                        var tableItemPattern = element.GetCurrentPattern(TableItemPattern.Pattern) as TableItemPattern;
                        if (tableItemPattern != null)
                        {
                            info.IsTableCell = true;
                            info.RowIndex = tableItemPattern.Current.Row;
                            info.ColumnIndex = tableItemPattern.Current.Column;
                            info.RowSpan = tableItemPattern.Current.RowSpan;
                            info.ColumnSpan = tableItemPattern.Current.ColumnSpan;

                            // Get the containing table
                            try
                            {
                                var containingGrid = tableItemPattern.Current.ContainingGrid;
                                if (containingGrid != null)
                                {
                                    info.TableName = containingGrid.Current.Name;
                                    if (string.IsNullOrEmpty(info.TableName))
                                        info.TableName = containingGrid.Current.AutomationId;
                                }
                            }
                            catch { }

                            // Get column headers for this cell
                            try
                            {
                                var columnHeaders = tableItemPattern.Current.GetColumnHeaderItems();
                                if (columnHeaders != null)
                                {
                                    foreach (var header in columnHeaders)
                                    {
                                        string headerText = header.Current.Name;
                                        if (!string.IsNullOrEmpty(headerText))
                                            info.ColumnHeaders.Add(headerText);
                                    }
                                }
                            }
                            catch { }

                            // Get row headers for this cell
                            try
                            {
                                var rowHeaders = tableItemPattern.Current.GetRowHeaderItems();
                                if (rowHeaders != null)
                                {
                                    foreach (var header in rowHeaders)
                                    {
                                        string headerText = header.Current.Name;
                                        if (!string.IsNullOrEmpty(headerText))
                                            info.RowHeaders.Add(headerText);
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    else if (pattern == GridItemPattern.Pattern)
                    {
                        var gridItemPattern = element.GetCurrentPattern(GridItemPattern.Pattern) as GridItemPattern;
                        if (gridItemPattern != null)
                        {
                            // Only set if not already set by TableItemPattern
                            if (info.RowIndex < 0)
                                info.RowIndex = gridItemPattern.Current.Row;
                            if (info.ColumnIndex < 0)
                                info.ColumnIndex = gridItemPattern.Current.Column;
                            if (info.RowSpan <= 1)
                                info.RowSpan = gridItemPattern.Current.RowSpan;
                            if (info.ColumnSpan <= 1)
                                info.ColumnSpan = gridItemPattern.Current.ColumnSpan;

                            info.IsTableCell = true;

                            // Get the containing grid
                            try
                            {
                                var containingGrid = gridItemPattern.Current.ContainingGrid;
                                if (containingGrid != null && string.IsNullOrEmpty(info.TableName))
                                {
                                    info.TableName = containingGrid.Current.Name;
                                    if (string.IsNullOrEmpty(info.TableName))
                                        info.TableName = containingGrid.Current.AutomationId;
                                }
                            }
                            catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Pattern extraction error: {ex.Message}");
            }
        }

        private void ExtractLegacyProperties(AutomationElement element, ElementInfo info)
        {
            try
            {
                // Try to get LegacyIAccessiblePattern - important for web elements
                // This pattern may not be available on all systems, so we'll try and catch
                try
                {
                    // LegacyIAccessiblePattern is in System.Windows.Automation but might not be available
                    // We'll use reflection to check if it's available
                    var patterns = element.GetSupportedPatterns();
                    foreach (var pattern in patterns)
                    {
                        if (pattern.ProgrammaticName.Contains("LegacyIAccessible"))
                        {
                            Debug.WriteLine($"LegacyIAccessiblePattern is supported for this element");
                            // Try to get the pattern dynamically
                            try
                            {
                                var legacyPattern = element.GetCurrentPattern(pattern);
                                if (legacyPattern != null)
                                {
                                    // Use reflection to get properties
                                    var currentProp = legacyPattern.GetType().GetProperty("Current");
                                    if (currentProp != null)
                                    {
                                        var current = currentProp.GetValue(legacyPattern);
                                        if (current != null)
                                        {
                                            var currentType = current.GetType();

                                            // Extract all legacy properties using reflection
                                            try
                                            {
                                                var nameProp = currentType.GetProperty("Name");
                                                info.LegacyName = nameProp?.GetValue(current)?.ToString();
                                            }
                                            catch { }

                                            try
                                            {
                                                var valueProp = currentType.GetProperty("Value");
                                                info.LegacyValue = valueProp?.GetValue(current)?.ToString();
                                            }
                                            catch { }

                                            try
                                            {
                                                var descProp = currentType.GetProperty("Description");
                                                info.LegacyDescription = descProp?.GetValue(current)?.ToString();
                                            }
                                            catch { }

                                            try
                                            {
                                                var helpProp = currentType.GetProperty("Help");
                                                info.LegacyHelp = helpProp?.GetValue(current)?.ToString();
                                            }
                                            catch { }

                                            try
                                            {
                                                var keyProp = currentType.GetProperty("KeyboardShortcut");
                                                info.LegacyKeyboardShortcut = keyProp?.GetValue(current)?.ToString();
                                            }
                                            catch { }

                                            try
                                            {
                                                var defaultActionProp = currentType.GetProperty("DefaultAction");
                                                info.LegacyDefaultAction = defaultActionProp?.GetValue(current)?.ToString();
                                            }
                                            catch { }

                                            try
                                            {
                                                var stateProp = currentType.GetProperty("State");
                                                var stateValue = stateProp?.GetValue(current);
                                                if (stateValue != null && stateValue is int)
                                                    info.LegacyState = (int)stateValue;
                                            }
                                            catch { }

                                            try
                                            {
                                                var roleProp = currentType.GetProperty("Role");
                                                var roleValue = roleProp?.GetValue(current);
                                                if (roleValue != null && roleValue is int)
                                                    info.LegacyRole = (int)roleValue;
                                            }
                                            catch { }

                                            try
                                            {
                                                var childCountProp = currentType.GetProperty("ChildCount");
                                                var childCountValue = childCountProp?.GetValue(current);
                                                if (childCountValue != null && childCountValue is int)
                                                    info.LegacyChildCount = (int)childCountValue;
                                            }
                                            catch { }

                                            Debug.WriteLine($"LegacyPattern extracted - Name: '{info.LegacyName}', Value: '{info.LegacyValue}', Role: {info.LegacyRole}, State: {info.LegacyState}, DefaultAction: '{info.LegacyDefaultAction}', ChildCount: {info.LegacyChildCount}");
                                        }
                                    }
                                }
                            }
                            catch (Exception innerEx)
                            {
                                Debug.WriteLine($"Error accessing LegacyIAccessiblePattern properties: {innerEx.Message}");
                            }
                            break;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"LegacyIAccessiblePattern not available: {ex.Message}");
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Legacy properties error: {ex.Message}");
            }
        }

        private void BuildWindowsPath(AutomationElement element, ElementInfo info)
        {
            try
            {
                var path = new List<string>();
                var treePath = new List<string>();
                var elementPath = new List<string>();
                var current = element;
                int childIndex = 0;

                while (current != null && current != AutomationElement.RootElement)
                {
                    var name = current.Current.Name;
                    var className = current.Current.ClassName;
                    var id = current.Current.AutomationId;
                    var controlType = current.Current.ControlType?.ProgrammaticName ?? "Unknown";

                    // Windows Path (simple, readable)
                    var pathPart = !string.IsNullOrEmpty(name) ? name :
                                   !string.IsNullOrEmpty(id) ? $"#{id}" :
                                   !string.IsNullOrEmpty(className) ? $".{className}" :
                                   "Unknown";

                    // Tree Path (includes control type)
                    var treePathPart = $"{controlType}";
                    if (!string.IsNullOrEmpty(name))
                        treePathPart += $"[Name='{name}']";
                    else if (!string.IsNullOrEmpty(id))
                        treePathPart += $"[AutomationId='{id}']";

                    // Element Path (full details with index)
                    var elemPathPart = $"{controlType}[{childIndex}]";
                    if (!string.IsNullOrEmpty(name))
                        elemPathPart += $"{{Name='{name}'}}";

                    path.Insert(0, pathPart);
                    treePath.Insert(0, treePathPart);
                    elementPath.Insert(0, elemPathPart);

                    // Get child index for next iteration
                    try
                    {
                        var parent = TreeWalker.RawViewWalker.GetParent(current);
                        if (parent != null)
                        {
                            var children = parent.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
                            childIndex = 0;
                            foreach (AutomationElement child in children)
                            {
                                if (child == current)
                                    break;
                                childIndex++;
                            }
                        }
                    }
                    catch { }

                    current = TreeWalker.RawViewWalker.GetParent(current);
                }

                info.WindowsPath = string.Join(" > ", path);
                info.TreePath = string.Join(" / ", treePath);
                info.ElementPath = string.Join(" / ", elementPath);

                // Also generate advanced Windows path using SelectorGenerator
                if (string.IsNullOrEmpty(info.WindowsPath) || info.WindowsPath == "")
                {
                    info.WindowsPath = Utils.SelectorGenerator.GenerateWindowsPath(element);
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Windows path error: {ex.Message}");
            }
        }

        private void CollectAllElements(AutomationElement parent, List<ElementInfo> elements, CollectionProfile profile, int level)
        {
            try
            {
                var info = ExtractElementInfo(parent, profile);
                if (info != null)
                {
                    info.TreeLevel = level;
                    elements.Add(info);
                }

                var children = parent.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
                foreach (AutomationElement child in children)
                {
                    CollectAllElements(child, elements, profile, level + 1);
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"CollectAllElements error at level {level}: {ex.Message}");
            }
        }

        private ElementInfo BuildElementTree(AutomationElement element, CollectionProfile profile, int level)
        {
            var info = ExtractElementInfo(element, profile);
            if (info == null) return null;

            info.TreeLevel = level;

            try
            {
                var children = element.FindAll(TreeScope.Children, System.Windows.Automation.Condition.TrueCondition);
                foreach (AutomationElement child in children)
                {
                    var childInfo = BuildElementTree(child, profile, level + 1);
                    if (childInfo != null)
                    {
                        childInfo.Parent = info;
                        info.Children.Add(childInfo);
                    }
                }
            }
            catch (Exception ex)
            {
                info.CollectionErrors.Add($"Tree building error: {ex.Message}");
            }

            return info;
        }
    }
}