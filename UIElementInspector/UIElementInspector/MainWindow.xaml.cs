using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Threading;
using UIElementInspector.Core.Detectors;
using UIElementInspector.Core.Models;
using UIElementInspector.Core.Utils;
using UIElementInspector.Services;
using Newtonsoft.Json;

namespace UIElementInspector
{
    /// <summary>
    /// Main application window for Universal UI Element Inspector
    /// </summary>
    public partial class MainWindow : Window
    {
        private readonly List<IElementDetector> _detectors;
        private readonly ObservableCollection<ElementInfo> _collectedElements;
        private MouseHookService _mouseHook;
        private HotkeyService _hotkeyService;
        private ElementInfo _currentElement;
        private bool _isInspecting;
        private CancellationTokenSource _inspectionCts;
        private DispatcherTimer _memoryTimer;
        private DispatcherTimer _mouseTimer;

        public MainWindow()
        {
            InitializeComponent();

            // Initialize collections
            _detectors = new List<IElementDetector>();
            _collectedElements = new ObservableCollection<ElementInfo>();

            // Bind tree view to collection
            tvElements.ItemsSource = _collectedElements;

            // Initialize services
            InitializeServices();

            // Initialize detectors
            InitializeDetectors();

            // Setup timers
            SetupTimers();

            // Log startup
            LogToConsole("Universal UI Element Inspector started successfully.");
            LogToConsole($"Available detectors: {string.Join(", ", _detectors.Select(d => d.Name))}");
        }

        private void InitializeServices()
        {
            try
            {
                // Initialize mouse hook service
                _mouseHook = new MouseHookService();
                _mouseHook.MouseMove += OnGlobalMouseMove;
                _mouseHook.MouseClick += OnGlobalMouseClick;

                // Initialize hotkey service
                _hotkeyService = new HotkeyService(this);
                _hotkeyService.RegisterHotkey(Key.F1, ModifierKeys.None, StartInspection_Click);
                _hotkeyService.RegisterHotkey(Key.Escape, ModifierKeys.None, StopInspection_Click);
                _hotkeyService.RegisterHotkey(Key.F5, ModifierKeys.None, Refresh_Click);
                _hotkeyService.RegisterHotkey(Key.S, ModifierKeys.Control, ExportQuick_Click);

                LogToConsole("Services initialized successfully.");
            }
            catch (Exception ex)
            {
                LogToConsole($"Error initializing services: {ex.Message}");
            }
        }

        private void InitializeDetectors()
        {
            try
            {
                // Add UI Automation detector
                _detectors.Add(new UIAutomationDetector());

                // Add WebView2/CDP detector for modern web
                _detectors.Add(new WebView2Detector());

                // Add MSHTML detector for IE and embedded browsers
                _detectors.Add(new MSHTMLDetector());

                // TODO: Add other detectors as they are implemented
                // _detectors.Add(new PlaywrightDetector());

                LogToConsole($"Initialized {_detectors.Count} detector(s).");
            }
            catch (Exception ex)
            {
                LogToConsole($"Error initializing detectors: {ex.Message}");
            }
        }

        private void SetupTimers()
        {
            // Memory usage timer
            _memoryTimer = new DispatcherTimer();
            _memoryTimer.Interval = TimeSpan.FromSeconds(2);
            _memoryTimer.Tick += (s, e) =>
            {
                var process = Process.GetCurrentProcess();
                var memoryMB = process.WorkingSet64 / (1024 * 1024);
                sbMemory.Text = $"Memory: {memoryMB} MB";
            };
            _memoryTimer.Start();

            // Mouse position timer
            _mouseTimer = new DispatcherTimer();
            _mouseTimer.Interval = TimeSpan.FromMilliseconds(100);
            _mouseTimer.Tick += (s, e) =>
            {
                var pos = System.Windows.Forms.Cursor.Position;
                sbMousePosition.Text = $"Mouse: {pos.X}, {pos.Y}";
            };
            _mouseTimer.Start();
        }

        #region Inspection Methods

        private async void StartInspection_Click(object sender, RoutedEventArgs e)
        {
            if (_isInspecting)
            {
                LogToConsole("Inspection is already running.");
                return;
            }

            _isInspecting = true;
            _inspectionCts = new CancellationTokenSource();

            // Update UI
            btnStartInspection.IsEnabled = false;
            btnStopInspection.IsEnabled = true;
            txtStatus.Text = "Inspecting...";
            txtStatus.Foreground = System.Windows.Media.Brushes.Orange;
            sbStatus.Text = "Inspection mode active";

            // Clear previous data
            _collectedElements.Clear();
            ClearElementDetails();

            // Get inspection mode
            var mode = cmbInspectionMode.SelectedIndex;

            LogToConsole($"Starting inspection in mode: {((ComboBoxItem)cmbInspectionMode.SelectedItem).Content}");

            try
            {
                switch (mode)
                {
                    case 0: // Hover (Real-time)
                        _mouseHook.StartHook();
                        break;

                    case 1: // Click (Snapshot)
                        _mouseHook.StartHook();
                        LogToConsole("Click on an element to capture it.");
                        break;

                    case 2: // Region Select
                        await StartRegionSelection();
                        break;

                    case 3: // Full Window
                        await CaptureFullWindow();
                        break;
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Error starting inspection: {ex.Message}");
                StopInspection_Click(null, null);
            }
        }

        private void StopInspection_Click(object sender, RoutedEventArgs e)
        {
            if (!_isInspecting) return;

            _isInspecting = false;
            _inspectionCts?.Cancel();

            // Update UI
            btnStartInspection.IsEnabled = true;
            btnStopInspection.IsEnabled = false;
            txtStatus.Text = "Ready";
            txtStatus.Foreground = System.Windows.Media.Brushes.Green;
            sbStatus.Text = "Ready";

            // Stop mouse hook
            _mouseHook.StopHook();

            LogToConsole("Inspection stopped.");
            LogToConsole($"Total elements collected: {_collectedElements.Count}");
        }

        private async void OnGlobalMouseMove(object sender, System.Windows.Point point)
        {
            if (!_isInspecting) return;

            var mode = cmbInspectionMode.SelectedIndex;
            if (mode != 0) return; // Only for hover mode

            await CaptureElementAtPoint(point);
        }

        private async void OnGlobalMouseClick(object sender, System.Windows.Point point)
        {
            if (!_isInspecting) return;

            var mode = cmbInspectionMode.SelectedIndex;
            if (mode != 1) return; // Only for click mode

            await CaptureElementAtPoint(point);

            // Stop inspection after click in snapshot mode
            if (mode == 1)
            {
                StopInspection_Click(null, null);
            }
        }

        private async Task CaptureElementAtPoint(System.Windows.Point point)
        {
            try
            {
                var profile = GetSelectedProfile();
                var detector = GetSelectedDetector();

                if (detector == null)
                {
                    LogToConsole("No detector available for the current target.");
                    return;
                }

                var stopwatch = Stopwatch.StartNew();
                var element = await detector.GetElementAtPoint(point, profile);
                stopwatch.Stop();

                if (element != null)
                {
                    _currentElement = element;

                    // Update UI on UI thread
                    await Dispatcher.InvokeAsync(() =>
                    {
                        DisplayElementInfo(element);
                        txtCollectionTime.Text = $"{stopwatch.ElapsedMilliseconds} ms";
                        txtDetectionMethod.Text = element.DetectionMethod;

                        // Add to collection if not already present
                        if (!_collectedElements.Any(e => e.Id == element.Id))
                        {
                            _collectedElements.Add(element);
                            txtElementCount.Text = _collectedElements.Count.ToString();
                        }

                        sbElementInfo.Text = $"{element.ElementType}: {element.Name ?? "Unnamed"}";
                    });
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Error capturing element: {ex.Message}");
            }
        }

        private async Task StartRegionSelection()
        {
            LogToConsole("Starting region selection mode. Draw a rectangle with mouse.");

            try
            {
                // Hide main window temporarily
                this.Visibility = Visibility.Hidden;
                await Task.Delay(100); // Let the window fully hide

                // Show region selector
                var selectedRegion = Windows.RegionSelectorWindow.SelectRegion();

                // Show main window again
                this.Visibility = Visibility.Visible;

                if (selectedRegion.HasValue)
                {
                    var region = selectedRegion.Value;
                    LogToConsole($"Region selected: X={region.X:F0}, Y={region.Y:F0}, Width={region.Width:F0}, Height={region.Height:F0}");

                    // Get elements in the selected region
                    ShowProgress(true, "Collecting elements in selected region...");

                    var profile = GetSelectedProfile();
                    var detector = GetSelectedDetector();

                    if (detector != null)
                    {
                        var elements = await detector.GetElementsInRegion(region, profile);

                        await Dispatcher.InvokeAsync(() =>
                        {
                            _collectedElements.Clear();
                            foreach (var element in elements)
                            {
                                _collectedElements.Add(element);
                            }

                            txtElementCount.Text = elements.Count.ToString();
                            LogToConsole($"Found {elements.Count} elements in the selected region");

                            // Optionally take a screenshot of the region
                            if (System.Windows.MessageBox.Show("Do you want to take a screenshot of the selected region?",
                                "Screenshot", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                            {
                                var screenshot = Core.Utils.ScreenshotHelper.CaptureRegion(
                                    new System.Drawing.Rectangle(
                                        (int)region.X, (int)region.Y,
                                        (int)region.Width, (int)region.Height));

                                if (screenshot != null)
                                {
                                    // Save to desktop
                                    var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                                    var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                                    var fileName = $"Region_Screenshot_{timestamp}.png";
                                    var filePath = System.IO.Path.Combine(desktop, fileName);

                                    Core.Utils.ScreenshotHelper.SaveToFile(screenshot, filePath);
                                    LogToConsole($"Region screenshot saved to: {filePath}");

                                    // Display in UI
                                    var bitmapImage = Core.Utils.ScreenshotHelper.ConvertToBitmapImage(screenshot);
                                    imgScreenshot.Source = bitmapImage;

                                    screenshot.Dispose();
                                }
                            }
                        });
                    }
                    else
                    {
                        LogToConsole("No detector available for region selection.");
                    }

                    ShowProgress(false);
                }
                else
                {
                    LogToConsole("Region selection cancelled.");
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Error during region selection: {ex.Message}");
                this.Visibility = Visibility.Visible;
                ShowProgress(false);
            }
            finally
            {
                StopInspection_Click(null, null);
            }
        }

        private async Task CaptureFullWindow()
        {
            try
            {
                ShowProgress(true, "Capturing all elements...");

                var profile = GetSelectedProfile();
                var detector = GetSelectedDetector();

                if (detector == null)
                {
                    LogToConsole("No detector available.");
                    return;
                }

                var stopwatch = Stopwatch.StartNew();
                var elements = await detector.GetAllElements(IntPtr.Zero, profile);
                stopwatch.Stop();

                await Dispatcher.InvokeAsync(() =>
                {
                    _collectedElements.Clear();
                    foreach (var element in elements)
                    {
                        _collectedElements.Add(element);
                    }

                    txtElementCount.Text = elements.Count.ToString();
                    txtCollectionTime.Text = $"{stopwatch.ElapsedMilliseconds} ms";

                    LogToConsole($"Captured {elements.Count} elements in {stopwatch.ElapsedMilliseconds}ms");
                });
            }
            catch (Exception ex)
            {
                LogToConsole($"Error capturing full window: {ex.Message}");
            }
            finally
            {
                ShowProgress(false);
                StopInspection_Click(null, null);
            }
        }

        #endregion

        #region Display Methods

        private void DisplayElementInfo(ElementInfo element)
        {
            if (element == null)
            {
                ClearElementDetails();
                return;
            }

            // Display raw properties
            txtRawProperties.Text = GenerateRawProperties(element);

            // Display categorized properties
            DisplayCategorizedProperties(element);

            // Display selectors - Generate if not available
            if (string.IsNullOrEmpty(element.XPath))
            {
                element.XPath = SelectorGenerator.GetOptimalXPath(element);
            }
            txtXPath.Text = element.XPath ?? "N/A";

            if (string.IsNullOrEmpty(element.CssSelector))
            {
                element.CssSelector = SelectorGenerator.GetOptimalCssSelector(element);
            }
            txtCssSelector.Text = element.CssSelector ?? "N/A";

            txtWindowsPath.Text = element.WindowsPath ?? "N/A";

            // Show alternative selectors in tooltip
            var xpathStrategies = SelectorGenerator.GenerateXPathStrategies(element);
            if (xpathStrategies.Count > 1)
            {
                txtXPath.ToolTip = "Alternative XPaths:\n" + string.Join("\n", xpathStrategies.Take(5));
            }

            var cssStrategies = SelectorGenerator.GenerateCssSelectorStrategies(element);
            if (cssStrategies.Count > 1)
            {
                txtCssSelector.ToolTip = "Alternative CSS Selectors:\n" + string.Join("\n", cssStrategies.Take(5));
            }

            // Display source code if available
            txtSourceCode.Text = element.SourceCode ?? "No source code available";

            // Display screenshot if available
            if (element.Screenshot != null && element.Screenshot.Length > 0)
            {
                var bitmap = new BitmapImage();
                bitmap.BeginInit();
                bitmap.StreamSource = new MemoryStream(element.Screenshot);
                bitmap.EndInit();
                imgScreenshot.Source = bitmap;
            }
            else
            {
                imgScreenshot.Source = null;
            }
        }

        private string GenerateRawProperties(ElementInfo element)
        {
            var sb = new StringBuilder();
            sb.AppendLine("=== Element Information ===");
            sb.AppendLine($"Capture Time: {element.CaptureTime}");
            sb.AppendLine($"Detection Method: {element.DetectionMethod}");
            sb.AppendLine($"Collection Profile: {element.CollectionProfile}");
            sb.AppendLine($"Collection Duration: {element.CollectionDuration.TotalMilliseconds}ms");
            sb.AppendLine();

            sb.AppendLine("=== Basic Properties ===");
            AppendProperty(sb, "Element Type", element.ElementType);
            AppendProperty(sb, "Name", element.Name);
            AppendProperty(sb, "Class Name", element.ClassName);
            AppendProperty(sb, "Value", element.Value);
            AppendProperty(sb, "Description", element.Description);
            sb.AppendLine();

            sb.AppendLine("=== UI Automation Properties ===");
            AppendProperty(sb, "AutomationId", element.AutomationId);
            AppendProperty(sb, "Control Type", element.ControlType);
            AppendProperty(sb, "Localized Control Type", element.LocalizedControlType);
            AppendProperty(sb, "Framework Id", element.FrameworkId);
            AppendProperty(sb, "Process Id", element.ProcessId.ToString());
            AppendProperty(sb, "Runtime Id", element.RuntimeId);
            AppendProperty(sb, "Native Window Handle", element.NativeWindowHandle.ToString());
            AppendProperty(sb, "Is Enabled", element.IsEnabled.ToString());
            AppendProperty(sb, "Is Offscreen", element.IsOffscreen.ToString());
            AppendProperty(sb, "Has Keyboard Focus", element.HasKeyboardFocus.ToString());
            AppendProperty(sb, "Is Keyboard Focusable", element.IsKeyboardFocusable.ToString());
            AppendProperty(sb, "Is Password", element.IsPassword.ToString());
            AppendProperty(sb, "Help Text", element.HelpText);
            AppendProperty(sb, "Accelerator Key", element.AcceleratorKey);
            AppendProperty(sb, "Access Key", element.AccessKey);

            if (element.SupportedPatterns?.Count > 0)
            {
                sb.AppendLine($"Supported Patterns: {string.Join(", ", element.SupportedPatterns)}");
            }
            sb.AppendLine();

            sb.AppendLine("=== Position & Size ===");
            sb.AppendLine($"Bounding Rectangle: {element.BoundingRectangle}");
            sb.AppendLine($"X: {element.X}, Y: {element.Y}");
            sb.AppendLine($"Width: {element.Width}, Height: {element.Height}");
            sb.AppendLine($"Clickable Point: {element.ClickablePoint}");
            sb.AppendLine();

            if (!string.IsNullOrEmpty(element.TagName))
            {
                sb.AppendLine("=== Web/HTML Properties ===");
                AppendProperty(sb, "Tag Name", element.TagName);
                AppendProperty(sb, "HTML Id", element.HtmlId);
                AppendProperty(sb, "HTML Class", element.HtmlClassName);
                AppendProperty(sb, "Inner Text", element.InnerText);
                AppendProperty(sb, "Href", element.Href);
                AppendProperty(sb, "Role", element.Role);
                AppendProperty(sb, "Aria Label", element.AriaLabel);
                sb.AppendLine();
            }

            sb.AppendLine("=== Hierarchy ===");
            AppendProperty(sb, "Parent Name", element.ParentName);
            AppendProperty(sb, "Parent Id", element.ParentId);
            AppendProperty(sb, "Parent Class", element.ParentClassName);
            sb.AppendLine($"Children Count: {element.Children?.Count ?? 0}");
            sb.AppendLine($"Tree Level: {element.TreeLevel}");
            sb.AppendLine();

            if (element.CustomProperties?.Count > 0)
            {
                sb.AppendLine("=== Custom Properties ===");
                foreach (var prop in element.CustomProperties)
                {
                    sb.AppendLine($"{prop.Key}: {prop.Value}");
                }
                sb.AppendLine();
            }

            if (element.CollectionErrors?.Count > 0)
            {
                sb.AppendLine("=== Collection Errors ===");
                foreach (var error in element.CollectionErrors)
                {
                    sb.AppendLine($"- {error}");
                }
            }

            return sb.ToString();
        }

        private void AppendProperty(StringBuilder sb, string name, string value)
        {
            if (!string.IsNullOrEmpty(value))
            {
                sb.AppendLine($"{name}: {value}");
            }
        }

        private void DisplayCategorizedProperties(ElementInfo element)
        {
            // Basic properties
            var basicProps = new List<KeyValuePair<string, string>>
            {
                new("Element Type", element.ElementType ?? ""),
                new("Name", element.Name ?? ""),
                new("Class Name", element.ClassName ?? ""),
                new("Value", element.Value ?? ""),
                new("Description", element.Description ?? "")
            };
            dgBasicProperties.ItemsSource = basicProps.Where(p => !string.IsNullOrEmpty(p.Value));

            // UI Automation properties
            var uiaProps = new List<KeyValuePair<string, string>>
            {
                new("AutomationId", element.AutomationId ?? ""),
                new("Control Type", element.ControlType ?? ""),
                new("Framework Id", element.FrameworkId ?? ""),
                new("Process Id", element.ProcessId.ToString()),
                new("Is Enabled", element.IsEnabled.ToString()),
                new("Is Offscreen", element.IsOffscreen.ToString())
            };
            dgUIAutomationProperties.ItemsSource = uiaProps.Where(p => !string.IsNullOrEmpty(p.Value));

            // Web properties
            var webProps = new List<KeyValuePair<string, string>>
            {
                new("Tag Name", element.TagName ?? ""),
                new("HTML Id", element.HtmlId ?? ""),
                new("HTML Class", element.HtmlClassName ?? ""),
                new("Inner Text", element.InnerText ?? ""),
                new("Role", element.Role ?? ""),
                new("Aria Label", element.AriaLabel ?? "")
            };
            dgWebProperties.ItemsSource = webProps.Where(p => !string.IsNullOrEmpty(p.Value));
        }

        private void ClearElementDetails()
        {
            txtRawProperties.Clear();
            dgBasicProperties.ItemsSource = null;
            dgUIAutomationProperties.ItemsSource = null;
            dgWebProperties.ItemsSource = null;
            txtXPath.Clear();
            txtCssSelector.Clear();
            txtWindowsPath.Clear();
            txtSourceCode.Clear();
            imgScreenshot.Source = null;
        }

        #endregion

        #region Helper Methods

        private CollectionProfile GetSelectedProfile()
        {
            return (CollectionProfile)cmbCollectionProfile.SelectedIndex;
        }

        private IElementDetector GetSelectedDetector()
        {
            var techIndex = cmbDetectionTech.SelectedIndex;

            if (techIndex == 0) // Auto Detect
            {
                // Return first detector that can detect current point
                var point = System.Windows.Forms.Cursor.Position;
                var wpfPoint = new System.Windows.Point(point.X, point.Y);
                return _detectors.FirstOrDefault(d => d.CanDetect(wpfPoint));
            }
            else if (techIndex == 1) // UI Automation
            {
                return _detectors.FirstOrDefault(d => d.Name == "UI Automation");
            }
            else if (techIndex == 2) // WebView2/CDP
            {
                return _detectors.FirstOrDefault(d => d.Name == "WebView2/CDP");
            }
            else if (techIndex == 3) // MSHTML
            {
                return _detectors.FirstOrDefault(d => d.Name == "MSHTML");
            }
            else if (techIndex == 5) // All Technologies
            {
                // Return the first detector that can detect
                var point = System.Windows.Forms.Cursor.Position;
                var wpfPoint = new System.Windows.Point(point.X, point.Y);
                return _detectors.FirstOrDefault(d => d.CanDetect(wpfPoint)) ?? _detectors.FirstOrDefault();
            }
            // TODO: Add other detectors as they are implemented

            return _detectors.FirstOrDefault();
        }

        private void LogToConsole(string message)
        {
            Dispatcher.Invoke(() =>
            {
                var timestamp = DateTime.Now.ToString("HH:mm:ss");
                txtConsole.AppendText($"[{timestamp}] {message}\n");
                txtConsole.ScrollToEnd();
            });
        }

        private void ShowProgress(bool show, string message = "")
        {
            Dispatcher.Invoke(() =>
            {
                pbProgress.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
                txtProgress.Visibility = show ? Visibility.Visible : Visibility.Collapsed;
                txtProgress.Text = message;

                if (show)
                {
                    pbProgress.IsIndeterminate = true;
                }
            });
        }

        #endregion

        #region Event Handlers

        private void TreeView_SelectedItemChanged(object sender, RoutedPropertyChangedEventArgs<object> e)
        {
            var selectedElement = tvElements.SelectedItem as ElementInfo;
            if (selectedElement != null)
            {
                DisplayElementInfo(selectedElement);
                _currentElement = selectedElement;
            }
        }

        private void TreeView_MouseRightButtonUp(object sender, MouseButtonEventArgs e)
        {
            // TODO: Implement context menu for tree view
        }

        private void TreeSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            // TODO: Implement tree search functionality
        }

        private void TreeSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            if (txtTreeSearch.Text == "Search elements...")
            {
                txtTreeSearch.Text = "";
                txtTreeSearch.Foreground = System.Windows.Media.Brushes.Black;
            }
        }

        private void TreeSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTreeSearch.Text))
            {
                txtTreeSearch.Text = "Search elements...";
                txtTreeSearch.Foreground = System.Windows.Media.Brushes.Gray;
            }
        }

        private void TreeSearch_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement tree search
        }

        private void ExpandAll_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement expand all functionality
        }

        private void CollapseAll_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement collapse all functionality
        }

        private void Refresh_Click(object sender, RoutedEventArgs e)
        {
            if (_currentElement != null)
            {
                LogToConsole("Refreshing current element...");
                // TODO: Implement refresh functionality
            }
        }

        private void Screenshot_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_currentElement != null && _currentElement.BoundingRectangle != Rect.Empty)
                {
                    // Take screenshot of current element
                    var screenshot = Core.Utils.ScreenshotHelper.CaptureElement(_currentElement.BoundingRectangle);

                    if (screenshot != null)
                    {
                        // Convert to byte array and store in element
                        _currentElement.Screenshot = Core.Utils.ScreenshotHelper.ConvertToByteArray(screenshot);

                        // Display in UI
                        var bitmapImage = Core.Utils.ScreenshotHelper.ConvertToBitmapImage(screenshot);
                        imgScreenshot.Source = bitmapImage;

                        // Save to desktop with timestamp
                        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                        var fileName = $"Element_Screenshot_{timestamp}.png";
                        var filePath = System.IO.Path.Combine(desktop, fileName);

                        Core.Utils.ScreenshotHelper.SaveToFile(screenshot, filePath);

                        LogToConsole($"Screenshot saved to: {filePath}");

                        // Highlight the captured region briefly
                        Core.Utils.ScreenshotHelper.HighlightRegion(_currentElement.BoundingRectangle, 500);

                        screenshot.Dispose();
                    }
                }
                else
                {
                    LogToConsole("No element selected. Taking full screen screenshot...");

                    // Take full screen screenshot
                    var screenshot = Core.Utils.ScreenshotHelper.CaptureFullScreen();

                    if (screenshot != null)
                    {
                        // Save to desktop
                        var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                        var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                        var fileName = $"FullScreen_Screenshot_{timestamp}.png";
                        var filePath = System.IO.Path.Combine(desktop, fileName);

                        Core.Utils.ScreenshotHelper.SaveToFile(screenshot, filePath);

                        LogToConsole($"Full screen screenshot saved to: {filePath}");

                        screenshot.Dispose();
                    }
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Screenshot error: {ex.Message}");
                System.Windows.MessageBox.Show($"Failed to take screenshot: {ex.Message}", "Screenshot Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ExportQuick_Click(object sender, RoutedEventArgs e)
        {
            ExportToDesktop_Click(sender, e);
        }

        private void NewSession_Click(object sender, RoutedEventArgs e)
        {
            _collectedElements.Clear();
            ClearElementDetails();
            LogToConsole("New session started.");
        }

        private async void ExportCSV_Click(object sender, RoutedEventArgs e)
        {
            await ExportData("CSV");
        }

        private async void ExportTXT_Click(object sender, RoutedEventArgs e)
        {
            await ExportData("TXT");
        }

        private async void ExportJSON_Click(object sender, RoutedEventArgs e)
        {
            await ExportData("JSON");
        }

        private async void ExportXML_Click(object sender, RoutedEventArgs e)
        {
            await ExportData("XML");
        }

        private async void ExportHTML_Click(object sender, RoutedEventArgs e)
        {
            await ExportData("HTML");
        }

        private void ImportSession_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement import functionality
        }

        private void Exit_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void CopyElementData_Click(object sender, RoutedEventArgs e)
        {
            if (_currentElement != null)
            {
                var data = GenerateRawProperties(_currentElement);
                System.Windows.Clipboard.SetText(data);
                LogToConsole("Element data copied to clipboard.");
            }
        }

        private void CopyAllElements_Click(object sender, RoutedEventArgs e)
        {
            CopyAllElementsQuick_Click(sender, e);
        }

        private void CopyScreenshot_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement screenshot copy
        }

        private void CopySourceCode_Click(object sender, RoutedEventArgs e)
        {
            if (_currentElement?.SourceCode != null)
            {
                System.Windows.Clipboard.SetText(_currentElement.SourceCode);
                LogToConsole("Source code copied to clipboard.");
            }
        }

        private void ClearAll_Click(object sender, RoutedEventArgs e)
        {
            ClearAllData_Click(sender, e);
        }

        private void RawView_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Switch to raw view
        }

        private void TreeView_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Switch to tree view
        }

        private void ScreenshotTool_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Open screenshot tool
        }

        private void RegionSelector_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Open region selector
        }

        private void ColorPicker_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Open color picker
        }

        private void Settings_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Open settings window
        }

        private void Documentation_Click(object sender, RoutedEventArgs e)
        {
            Process.Start(new ProcessStartInfo
            {
                FileName = "https://github.com/yourusername/UIElementInspector",
                UseShellExecute = true
            });
        }

        private void KeyboardShortcuts_Click(object sender, RoutedEventArgs e)
        {
            var shortcuts = @"Keyboard Shortcuts:
F1 - Start Inspection
Esc - Stop Inspection
F5 - Refresh Current Element
Ctrl+S - Quick Export
Ctrl+C - Copy Element Data
Ctrl+Shift+C - Copy All Elements";

            System.Windows.MessageBox.Show(shortcuts, "Keyboard Shortcuts", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void About_Click(object sender, RoutedEventArgs e)
        {
            var about = @"Universal UI Element Inspector
Version 1.0.0

A comprehensive tool for collecting UI element data from any Windows application.

Supports multiple detection technologies:
- UI Automation
- WebView2/CDP
- MSHTML
- Playwright

© 2024 Your Company";

            System.Windows.MessageBox.Show(about, "About", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void CopyCurrentElement_Click(object sender, RoutedEventArgs e)
        {
            CopyElementData_Click(sender, e);
        }

        private void CopyAllElementsQuick_Click(object sender, RoutedEventArgs e)
        {
            if (_collectedElements.Count > 0)
            {
                var sb = new StringBuilder();
                foreach (var element in _collectedElements)
                {
                    sb.AppendLine(GenerateRawProperties(element));
                    sb.AppendLine("=" + new string('=', 50));
                }

                System.Windows.Clipboard.SetText(sb.ToString());
                LogToConsole($"Copied {_collectedElements.Count} elements to clipboard.");
            }
        }

        private void SaveScreenshot_Click(object sender, RoutedEventArgs e)
        {
            Screenshot_Click(sender, e); // Reuse the screenshot functionality
        }

        private async void ExportToDesktop_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
                var fileName = $"UIElements_{timestamp}.txt";
                var filePath = System.IO.Path.Combine(desktop, fileName);

                var sb = new StringBuilder();
                foreach (var element in _collectedElements)
                {
                    sb.AppendLine(GenerateRawProperties(element));
                    sb.AppendLine("=" + new string('=', 50));
                }

                await File.WriteAllTextAsync(filePath, sb.ToString());

                LogToConsole($"Exported to: {filePath}");
                System.Windows.MessageBox.Show($"Data exported to:\n{filePath}", "Export Successful",
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogToConsole($"Export error: {ex.Message}");
                System.Windows.MessageBox.Show($"Export failed: {ex.Message}", "Export Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void ClearAllData_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show("Clear all collected data?", "Confirm Clear",
                MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _collectedElements.Clear();
                ClearElementDetails();
                _currentElement = null;
                txtElementCount.Text = "0";
                txtCollectionTime.Text = "0 ms";
                txtDetectionMethod.Text = "None";
                LogToConsole("All data cleared.");
            }
        }

        private async Task ExportData(string format)
        {
            try
            {
                if (_collectedElements == null || _collectedElements.Count == 0)
                {
                    System.Windows.MessageBox.Show("No elements to export. Please collect some elements first.",
                        "No Data", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                // Create save file dialog
                var dialog = new Microsoft.Win32.SaveFileDialog();
                var timestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");

                switch (format.ToUpper())
                {
                    case "CSV":
                        dialog.Filter = "CSV Files (*.csv)|*.csv";
                        dialog.FileName = $"UIElements_{timestamp}.csv";
                        break;
                    case "TXT":
                        dialog.Filter = "Text Files (*.txt)|*.txt";
                        dialog.FileName = $"UIElements_{timestamp}.txt";
                        break;
                    case "JSON":
                        dialog.Filter = "JSON Files (*.json)|*.json";
                        dialog.FileName = $"UIElements_{timestamp}.json";
                        break;
                    case "XML":
                        dialog.Filter = "XML Files (*.xml)|*.xml";
                        dialog.FileName = $"UIElements_{timestamp}.xml";
                        break;
                    case "HTML":
                        dialog.Filter = "HTML Files (*.html)|*.html";
                        dialog.FileName = $"UIElements_{timestamp}.html";
                        break;
                    default:
                        LogToConsole($"Unsupported export format: {format}");
                        return;
                }

                if (dialog.ShowDialog() == true)
                {
                    ShowProgress(true, $"Exporting to {format}...");
                    LogToConsole($"Exporting {_collectedElements.Count} elements to {format}...");

                    var exportManager = new ExportManager();
                    string result = "";

                    await Task.Run(() =>
                    {
                        switch (format.ToUpper())
                        {
                            case "CSV":
                                result = exportManager.ExportToCSV(_collectedElements.ToList());
                                break;
                            case "TXT":
                                result = exportManager.ExportToText(_collectedElements.ToList());
                                break;
                            case "JSON":
                                result = exportManager.ExportToJSON(_collectedElements.ToList());
                                break;
                            case "XML":
                                result = exportManager.ExportToXML(_collectedElements.ToList());
                                break;
                            case "HTML":
                                result = exportManager.ExportToHTML(_collectedElements.ToList());
                                break;
                        }
                    });

                    // Save to file
                    await File.WriteAllTextAsync(dialog.FileName, result);

                    ShowProgress(false);
                    LogToConsole($"Successfully exported to: {dialog.FileName}");

                    // Ask if user wants to open the file
                    var openResult = System.Windows.MessageBox.Show(
                        $"Export successful!\n\nFile saved to:\n{dialog.FileName}\n\nDo you want to open the file?",
                        "Export Complete",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Question);

                    if (openResult == MessageBoxResult.Yes)
                    {
                        Process.Start(new ProcessStartInfo
                        {
                            FileName = dialog.FileName,
                            UseShellExecute = true
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                ShowProgress(false);
                LogToConsole($"Export error: {ex.Message}");
                System.Windows.MessageBox.Show($"Failed to export data: {ex.Message}",
                    "Export Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        #endregion

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            // Cleanup
            _mouseHook?.Dispose();
            _hotkeyService?.Dispose();
            _memoryTimer?.Stop();
            _mouseTimer?.Stop();
            _inspectionCts?.Cancel();
        }
    }
}