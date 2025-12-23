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
using UIElementInspector.Windows;
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
        private List<ElementInfo> _searchResults;
        private int _currentSearchIndex;
        private MouseHookService _mouseHook;
        private HotkeyService _hotkeyService;
        private ElementInfo _currentElement;
        private bool _isInspecting;
        private CancellationTokenSource _inspectionCts;
        private DispatcherTimer _memoryTimer;
        private DispatcherTimer _mouseTimer;
        private FloatingControlWindow _floatingWindow;
        private Core.Utils.Logger _logger;
        private ArchiveManager _archiveManager;

        public MainWindow()
        {
            InitializeComponent();

            // Initialize logger first
            _logger = new Core.Utils.Logger();
            _logger.LogSection("APPLICATION STARTUP");

            // Initialize collections
            _detectors = new List<IElementDetector>();
            _collectedElements = new ObservableCollection<ElementInfo>();

            // Bind tree view to collection
            tvElements.ItemsSource = _collectedElements;

            // Initialize floating control window
            InitializeFloatingWindow();

            // Initialize services
            InitializeServices();

            // Initialize detectors
            InitializeDetectors();

            // Setup timers
            SetupTimers();

            // Log startup
            LogToConsole("Universal UI Element Inspector started successfully.");
            LogToConsole($"Log file: {_logger.GetLogFilePath()}");
            LogToConsole($"Available detectors: {string.Join(", ", _detectors.Select(d => d.Name))}");
        }

        private void InitializeFloatingWindow()
        {
            _floatingWindow = new FloatingControlWindow();
            _floatingWindow.StopInspectionRequested += (s, e) => StopInspection_Click(s, new RoutedEventArgs());
            _floatingWindow.ShowMainWindowRequested += (s, e) => ShowMainWindow();
        }

        private void InitializeServices()
        {
            try
            {
                _logger.LogSection("SERVICE INITIALIZATION");

                // Initialize archive manager
                _archiveManager = new ArchiveManager();
                _logger.LogInfo($"Archive manager initialized - Path: {_archiveManager.ArchiveBasePath}");

                // Initialize mouse hook service
                _mouseHook = new MouseHookService();
                _mouseHook.MouseMove += OnGlobalMouseMove;
                _mouseHook.MouseClick += OnGlobalMouseClick;
                _logger.LogInfo("Mouse hook service initialized");

                // Initialize hotkey service
                _hotkeyService = new HotkeyService(this);

                // F1 = Start Inspection (pencere minimize edilir)
                var f1 = _hotkeyService.RegisterHotkey(Key.F1, ModifierKeys.None, StartInspection_Click);

                // F2 = Stop Inspection
                var f2 = _hotkeyService.RegisterHotkey(Key.F2, ModifierKeys.None, StopInspection_Click);

                // F3 = Start Inspection (pencere minimize edilmez - Keep Visible)
                var f3 = _hotkeyService.RegisterHotkey(Key.F3, ModifierKeys.None, StartKeepVisible_Click);

                // F4 = Shutter Mode - Deklansor (basili tutunca aktif, birakinca durur)
                _hotkeyService.RegisterShutterKey(Key.F4, ShutterDown, ShutterUp);

                // F5 = Refresh current element
                var f5 = _hotkeyService.RegisterHotkey(Key.F5, ModifierKeys.None, Refresh_Click);

                // F6 = Export all reports to Desktop AND Archive
                var f6 = _hotkeyService.RegisterHotkey(Key.F6, ModifierKeys.None, ExportToDesktopAndArchive_Click);

                // F7 = FULL CAPTURE - 5 teknoloji + element listesi + kaynak kod + screenshot
                var f7 = _hotkeyService.RegisterHotkey(Key.F7, ModifierKeys.None, FullCaptureToDesktopAndArchive_Click);

                // F8 = ARCHIVE ONLY - Save to archive folder only
                var f8 = _hotkeyService.RegisterHotkey(Key.F8, ModifierKeys.None, FullCaptureToArchiveOnly_Click);

                // Ctrl+S = Quick Export
                var ctrlS = _hotkeyService.RegisterHotkey(Key.S, ModifierKeys.Control, ExportQuick_Click);

                // F9 = Screenshot Region - Bolge secip ekran goruntusu al
                var f9 = _hotkeyService.RegisterHotkey(Key.F9, ModifierKeys.None, ScreenshotRegion_Click);

                _logger.LogInfo($"Hotkey service initialized - F1:{f1}, F2:{f2}, F3:{f3}, F4:Shutter, F5:{f5}, F6:{f6}, F7:{f7}, F8:{f8}, F9:{f9}, Ctrl+S:{ctrlS}");

                LogToConsole("===========================================");
                LogToConsole("          KISAYOL TUSLARI (HOTKEYS)        ");
                LogToConsole("===========================================");
                LogToConsole("  F1  = Start Inspection (Pencere Gizlenir)");
                LogToConsole("  F2  = Stop Inspection");
                LogToConsole("  F3  = Start Inspection (Pencere Gorunur)");
                LogToConsole("  F4  = DEKLANSOR (Basili Tut = Aktif)");
                LogToConsole("  F5  = Refresh Element");
                LogToConsole("  F6  = Masaustu + Arsiv (TXT Rapor)");
                LogToConsole("  F7  = TAM YAKALAMA (Masaustu + Arsiv)");
                LogToConsole("  F8  = SADECE ARSIV (Tam Yakalama)");
                LogToConsole("  F9  = EKRAN GORUNTUSU (Bolge Sec)");
                LogToConsole("  Ctrl+S = Hizli Export");
                LogToConsole("===========================================");

                // Initialize archive tab
                InitializeArchiveTab();
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, "Service initialization failed");
                LogToConsole($"Error initializing services: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        #region Shutter Mode (Deklansor)

        private bool _shutterActive = false;

        private void ShutterDown()
        {
            if (_shutterActive) return;
            _shutterActive = true;

            LogToConsole("[DEKLANSOR] Basili - Element yakalama AKTIF");
            txtStatus.Text = "DEKLANSOR AKTIF";
            txtStatus.Foreground = System.Windows.Media.Brushes.Red;

            // Update shutter status indicator
            txtShutterStatus.Text = ">>> F4 BASILI - YAKALAMA AKTIF <<<";
            brdShutterStatus.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(139, 0, 0));

            // Start capturing without minimizing window
            if (!_isInspecting)
            {
                _isInspecting = true;
                _inspectionCts = new CancellationTokenSource();
                _mouseHook.StartHook();
            }
        }

        private async void ShutterUp()
        {
            if (!_shutterActive) return;
            _shutterActive = false;

            LogToConsole("[DEKLANSOR] Birakildi - Element yakalaniyor...");

            // Show progress indicator with wait cursor
            StartProgress("Element yakalanıyor...", "Lütfen bekleyiniz, UI element bilgileri toplanıyor...");

            // Capture current element
            try
            {
                var point = System.Windows.Forms.Cursor.Position;
                var wpfPoint = new System.Windows.Point(point.X, point.Y);

                SetProgressValue(25, "Element bilgileri toplanıyor...", "Analiz ediliyor");
                await CaptureElementAtPoint(wpfPoint);
                SetProgressValue(75, "Element bilgileri alındı", "Tamamlanıyor");

                await Task.Delay(200); // Brief pause to show progress
                SetProgressValue(100, "Element yakalandı!", "Başarılı");
                await Task.Delay(300); // Show completion

                LogToConsole("[DEKLANSOR] Element basariyla yakalandi");
            }
            catch (Exception ex)
            {
                LogToConsole($"[DEKLANSOR] Yakalama hatasi: {ex.Message}", Core.Utils.LogLevel.Error);
            }

            // Stop progress and restore cursor
            StopProgress();

            // Stop inspection
            _isInspecting = false;
            _mouseHook.StopHook();
            _inspectionCts?.Cancel();

            txtStatus.Text = "Ready";
            txtStatus.Foreground = System.Windows.Media.Brushes.Green;

            // Clear shutter status indicator
            txtShutterStatus.Text = "";
            brdShutterStatus.Background = System.Windows.Media.Brushes.Transparent;
        }

        /// <summary>
        /// F4 Button Click - Shows info about shutter mode
        /// </summary>
        private void ShutterInfo_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.MessageBox.Show(
                "DEKLANŞÖR MODU (F4)\n\n" +
                "Kullanım:\n" +
                "1. F4 tuşuna BASILI TUTUN\n" +
                "2. Mouse'u yakalamak istediğiniz element üzerine götürün\n" +
                "3. F4 tuşunu BIRAKIN\n\n" +
                "Element otomatik olarak yakalanacaktır.\n\n" +
                "Not: Bu özellik klavye ile çalışır, butona tıklamak yerine F4 tuşunu kullanın.",
                "Deklanşör Modu - Bilgi",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }

        #endregion

        #region Desktop Export

        /// <summary>
        /// F6 - Export to both Desktop AND Archive (TXT Report)
        /// </summary>
        private void ExportToDesktopAndArchive_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_collectedElements.Count == 0)
                {
                    LogToConsole("Masaustune aktarilacak element yok!", Core.Utils.LogLevel.Warning);
                    System.Windows.MessageBox.Show("Aktarilacak element bulunamadi.\nOnce element yakalayin.", "Uyari", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var reportContent = GenerateQuickReportContent();

                // Create archive item
                var archiveItem = _archiveManager.CreateArchiveItem(
                    $"Quick Export {DateTime.Now:yyyy-MM-dd HH:mm}",
                    "QuickExport");

                // Save to archive
                var archiveFilePath = System.IO.Path.Combine(archiveItem.FolderPath, "Report.txt");
                System.IO.File.WriteAllText(archiveFilePath, reportContent, System.Text.Encoding.UTF8);
                archiveItem.FilePaths.Add(archiveFilePath);
                archiveItem.FileCount = 1;
                if (_collectedElements.Count > 0)
                {
                    archiveItem.ElementName = _collectedElements[0].Name;
                    archiveItem.ElementType = _collectedElements[0].ElementType;
                    archiveItem.WindowTitle = _collectedElements[0].WindowTitle;
                }
                _archiveManager.SaveIndex();

                // Save to Desktop
                var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                var desktopFilePath = System.IO.Path.Combine(desktopPath, $"UIElementInspector_Report_{timestamp}.txt");
                System.IO.File.WriteAllText(desktopFilePath, reportContent, System.Text.Encoding.UTF8);

                // Copy archive file path to clipboard
                System.Windows.Clipboard.SetText(archiveFilePath);

                LogToConsole($"RAPOR KAYDEDILDI:");
                LogToConsole($"  Arsiv: {archiveFilePath}");
                LogToConsole($"  Masaustu: {desktopFilePath}");
                LogToConsole($"  Dosya linki panoya kopyalandi!");

                RefreshArchiveList();

                System.Windows.MessageBox.Show(
                    $"Rapor hem arsive hem masaustune kaydedildi!\n\n" +
                    $"Arsiv: {archiveItem.FolderPath}\n" +
                    $"Masaustu: {desktopFilePath}\n\n" +
                    $"Dosya linki panoya kopyalandi!",
                    "Basarili", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogToConsole($"Kaydetme hatasi: {ex.Message}", Core.Utils.LogLevel.Error);
                System.Windows.MessageBox.Show($"Kaydetme hatasi:\n{ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Generate quick report content for F6
        /// </summary>
        private string GenerateQuickReportContent()
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("================================================================================");
            sb.AppendLine("                    UI ELEMENT INSPECTOR - RAPOR");
            sb.AppendLine($"                    Tarih: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"                    Toplam Element: {_collectedElements.Count}");
            sb.AppendLine("================================================================================");
            sb.AppendLine();

            // Kisayol listesi
            sb.AppendLine("KISAYOL TUSLARI:");
            sb.AppendLine("  F1  = Start Inspection (Pencere Gizlenir)");
            sb.AppendLine("  F2  = Stop Inspection");
            sb.AppendLine("  F3  = Start Inspection (Pencere Gorunur)");
            sb.AppendLine("  F4  = DEKLANSOR (Basili Tut = Aktif)");
            sb.AppendLine("  F5  = Refresh Element");
            sb.AppendLine("  F6  = Masaustu + Arsiv (TXT Rapor)");
            sb.AppendLine("  F7  = TAM YAKALAMA (Masaustu + Arsiv)");
            sb.AppendLine("  F8  = SADECE ARSIV (Tam Yakalama)");
            sb.AppendLine("  Ctrl+S = Hizli Export");
            sb.AppendLine();
            sb.AppendLine("================================================================================");
            sb.AppendLine();

            int index = 1;
            foreach (var element in _collectedElements)
            {
                sb.AppendLine($"--- ELEMENT {index} ---");
                sb.AppendLine($"Yakalama Zamani: {element.CaptureTime:yyyy-MM-dd HH:mm:ss}");
                sb.AppendLine($"Detection Method: {element.DetectionMethod}");
                sb.AppendLine($"Collection Profile: {element.CollectionProfile}");
                sb.AppendLine();

                // Temel Bilgiler
                sb.AppendLine("[TEMEL BILGILER]");
                if (!string.IsNullOrEmpty(element.Name)) sb.AppendLine($"  Name: {element.Name}");
                if (!string.IsNullOrEmpty(element.AutomationId)) sb.AppendLine($"  AutomationId: {element.AutomationId}");
                if (!string.IsNullOrEmpty(element.ClassName)) sb.AppendLine($"  ClassName: {element.ClassName}");
                if (!string.IsNullOrEmpty(element.ControlType)) sb.AppendLine($"  ControlType: {element.ControlType}");
                if (!string.IsNullOrEmpty(element.LocalizedControlType)) sb.AppendLine($"  LocalizedControlType: {element.LocalizedControlType}");
                sb.AppendLine();

                // Konum Bilgileri
                sb.AppendLine("[KONUM]");
                sb.AppendLine($"  X: {element.X}, Y: {element.Y}");
                sb.AppendLine($"  Width: {element.Width}, Height: {element.Height}");
                sb.AppendLine($"  BoundingRect: {element.BoundingRectangle}");
                sb.AppendLine();

                // Handle Bilgileri
                if (element.WindowHandle != IntPtr.Zero)
                {
                    sb.AppendLine("[HANDLE BILGILERI]");
                    sb.AppendLine($"  WindowHandle: 0x{element.WindowHandle.ToInt64():X}");
                    if (!string.IsNullOrEmpty(element.WindowTitle)) sb.AppendLine($"  WindowTitle: {element.WindowTitle}");
                    if (!string.IsNullOrEmpty(element.WindowClassName)) sb.AppendLine($"  WindowClassName: {element.WindowClassName}");
                    if (element.ProcessId > 0) sb.AppendLine($"  ProcessId: {element.ProcessId}");
                    sb.AppendLine();
                }

                // Web Bilgileri
                if (!string.IsNullOrEmpty(element.TagName) || !string.IsNullOrEmpty(element.HtmlId))
                {
                    sb.AppendLine("[WEB/HTML BILGILERI]");
                    if (!string.IsNullOrEmpty(element.TagName)) sb.AppendLine($"  TagName: {element.TagName}");
                    if (!string.IsNullOrEmpty(element.HtmlId)) sb.AppendLine($"  HTML Id: {element.HtmlId}");
                    if (!string.IsNullOrEmpty(element.HtmlClassName)) sb.AppendLine($"  HTML Class: {element.HtmlClassName}");
                    if (!string.IsNullOrEmpty(element.Href)) sb.AppendLine($"  Href: {element.Href}");
                    if (!string.IsNullOrEmpty(element.InnerText)) sb.AppendLine($"  InnerText: {element.InnerText.Substring(0, Math.Min(200, element.InnerText.Length))}...");
                    sb.AppendLine();
                }

                // XPath ve Selectors
                sb.AppendLine("[SELECTORS]");
                if (!string.IsNullOrEmpty(element.XPath)) sb.AppendLine($"  XPath: {element.XPath}");
                if (!string.IsNullOrEmpty(element.CssSelector)) sb.AppendLine($"  CSS Selector: {element.CssSelector}");
                if (!string.IsNullOrEmpty(element.PlaywrightSelector)) sb.AppendLine($"  Playwright: {element.PlaywrightSelector}");
                sb.AppendLine();

                // Durum Bilgileri
                sb.AppendLine("[DURUM]");
                sb.AppendLine($"  IsVisible: {element.IsVisible}");
                sb.AppendLine($"  IsEnabled: {element.IsEnabled}");
                sb.AppendLine($"  IsOffscreen: {element.IsOffscreen}");
                sb.AppendLine($"  HasKeyboardFocus: {element.HasKeyboardFocus}");
                sb.AppendLine();

                sb.AppendLine("--------------------------------------------------------------------------------");
                sb.AppendLine();
                index++;
            }

            return sb.ToString();
        }

        /// <summary>
        /// F7 - TAM YAKALAMA: Both Desktop AND Archive
        /// </summary>
        private async void FullCaptureToDesktopAndArchive_Click(object sender, RoutedEventArgs e)
        {
            await PerformFullCapture(saveToDesktop: true, saveToArchive: true);
        }

        /// <summary>
        /// F8 - TAM YAKALAMA: Only to Archive
        /// </summary>
        private async void FullCaptureToArchiveOnly_Click(object sender, RoutedEventArgs e)
        {
            await PerformFullCapture(saveToDesktop: false, saveToArchive: true);
        }

        /// <summary>
        /// F9 - Screenshot Region: Opens region selector, captures screenshot,
        /// saves to desktop, and copies both image and file path to clipboard
        /// </summary>
        private void ScreenshotRegion_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogToConsole("===========================================");
                LogToConsole("       EKRAN GORUNTUSU - BOLGE SEC         ");
                LogToConsole("===========================================");
                LogToConsole("Mouse ile bir bolge secin...");
                LogToConsole("ESC = Iptal");

                // Hide main window temporarily for cleaner screenshot
                var wasVisible = this.Visibility == Visibility.Visible;
                if (wasVisible)
                {
                    this.Hide();
                }

                // Small delay to let window hide
                System.Threading.Thread.Sleep(100);

                // Show region selector
                var selectedRegion = Windows.RegionSelectorWindow.SelectRegion();

                if (selectedRegion.HasValue && selectedRegion.Value.Width > 0 && selectedRegion.Value.Height > 0)
                {
                    var region = selectedRegion.Value;

                    // Convert WPF Rect to System.Drawing.Rectangle
                    var drawingRegion = new System.Drawing.Rectangle(
                        (int)region.X,
                        (int)region.Y,
                        (int)region.Width,
                        (int)region.Height);

                    // Capture, save to desktop, and copy to clipboard
                    var savedPath = Core.Utils.ScreenshotHelper.CaptureRegionToDesktopAndClipboard(drawingRegion);

                    // Show main window again
                    if (wasVisible)
                    {
                        this.Show();
                        this.Activate();
                    }

                    LogToConsole("-------------------------------------------");
                    LogToConsole("EKRAN GORUNTUSU BASARIYLA ALINDI!");
                    LogToConsole($"Konum: {savedPath}");
                    LogToConsole($"Boyut: {(int)region.Width} x {(int)region.Height} piksel");
                    LogToConsole("-------------------------------------------");
                    LogToConsole("PANOYA KOPYALANDI:");
                    LogToConsole("  - Resim (Paint, Word, vs. yapistir)");
                    LogToConsole("  - Dosya yolu (Notepad, vs. yapistir)");
                    LogToConsole("  - Dosya (Explorer'da yapistir)");
                    LogToConsole("===========================================");

                    SetOperationStatus("Screenshot alindi!", savedPath);
                }
                else
                {
                    // Show main window again if cancelled
                    if (wasVisible)
                    {
                        this.Show();
                        this.Activate();
                    }

                    LogToConsole("Ekran goruntusu iptal edildi.");
                    SetOperationStatus("Iptal edildi", "");
                }
            }
            catch (Exception ex)
            {
                // Make sure window is shown even on error
                this.Show();
                this.Activate();

                LogToConsole($"HATA: {ex.Message}");
                _logger.LogException(ex, "Screenshot region capture failed");
                SetOperationStatus("Hata!", ex.Message);
            }
        }

        /// <summary>
        /// Core full capture method - saves to specified locations with progress indicator
        /// </summary>
        private async Task PerformFullCapture(bool saveToDesktop, bool saveToArchive)
        {
            try
            {
                var targetDesc = saveToDesktop && saveToArchive ? "MASAUSTU + ARSIV" :
                                 saveToArchive ? "SADECE ARSIV" : "MASAUSTU";

                // Start progress with wait cursor
                SetOperationStatus("TAM YAKALAMA BASLIYOR...", "Adim 0/8", "Lütfen bekleyiniz, UI element bilgileri toplanıyor...");
                SetProgressValue(0, "TAM YAKALAMA BASLIYOR...", "Hazırlanıyor...");

                LogToConsole("===========================================");
                LogToConsole($"       TAM YAKALAMA ({targetDesc})      ");
                LogToConsole("===========================================");

                // Step 1: Capture element at point (0-10%)
                SetProgressValue(5, "Element bilgileri alınıyor...", "Adım 1/8");
                var mousePos = System.Windows.Forms.Cursor.Position;
                var wpfPoint = new System.Windows.Point(mousePos.X, mousePos.Y);
                await CaptureElementAtPoint(wpfPoint);
                LogToConsole($"Element bilgileri UI'a yuklendi");
                SetProgressValue(10, "Element bilgileri alındı", "Adım 1/8 tamamlandı");

                // Step 2: Create archive folder (10-15%)
                ArchiveItem archiveItem = null;
                string archiveFolderPath = null;
                if (saveToArchive)
                {
                    SetProgressValue(12, "Arşiv klasörü oluşturuluyor...", "Adım 2/8");
                    archiveItem = _archiveManager.CreateArchiveItem(
                        $"Full Capture {DateTime.Now:yyyy-MM-dd HH:mm}",
                        "FullCapture");
                    archiveFolderPath = archiveItem.FolderPath;
                    LogToConsole($"Arsiv klasoru olusturuldu: {archiveFolderPath}");
                }
                SetProgressValue(15, "Klasörler hazırlandı", "Adım 2/8 tamamlandı");

                // Create desktop folder if needed
                string desktopFolderPath = null;
                if (saveToDesktop)
                {
                    var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
                    var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    var folderName = $"UICapture_{timestamp}";
                    desktopFolderPath = System.IO.Path.Combine(desktopPath, folderName);
                    System.IO.Directory.CreateDirectory(desktopFolderPath);
                    LogToConsole($"Masaustu klasoru olusturuldu: {folderName}");
                }

                var savedFilePaths = new List<string>();
                int savedFiles = 0;

                // Step 3: 5 Technologies report (15-35%)
                SetProgressValue(20, "5 Teknoloji ile element bilgileri toplanıyor...", "Adım 3/8");
                LogToConsole("[3/8] 5 Teknoloji ile element bilgileri toplanıyor...");
                var allTechReport = await CaptureAllTechnologiesReport(wpfPoint);
                if (!string.IsNullOrEmpty(allTechReport))
                {
                    if (saveToArchive)
                    {
                        var archiveReportPath = System.IO.Path.Combine(archiveFolderPath, "01_Element_5Tech_Report.txt");
                        await System.IO.File.WriteAllTextAsync(archiveReportPath, allTechReport, System.Text.Encoding.UTF8);
                        savedFilePaths.Add(archiveReportPath);
                    }
                    if (saveToDesktop)
                    {
                        var desktopReportPath = System.IO.Path.Combine(desktopFolderPath, "01_Element_5Tech_Report.txt");
                        await System.IO.File.WriteAllTextAsync(desktopReportPath, allTechReport, System.Text.Encoding.UTF8);
                    }
                    LogToConsole($"  -> Element raporu kaydedildi");
                    savedFiles++;
                }
                SetProgressValue(35, "Element raporu tamamlandı", "Adım 3/8 tamamlandı");

                // Step 4: Page structure (35-50%)
                SetProgressValue(40, "Sayfa yapısı ve element listesi toplanıyor...", "Adım 4/8");
                LogToConsole("[4/8] Sayfa yapisi ve element listesi toplanıyor...");
                var pageStructureReport = await CapturePageStructureReport(wpfPoint);
                if (!string.IsNullOrEmpty(pageStructureReport))
                {
                    if (saveToArchive)
                    {
                        var archiveStructurePath = System.IO.Path.Combine(archiveFolderPath, "02_Page_Structure_ElementList.txt");
                        await System.IO.File.WriteAllTextAsync(archiveStructurePath, pageStructureReport, System.Text.Encoding.UTF8);
                        savedFilePaths.Add(archiveStructurePath);
                    }
                    if (saveToDesktop)
                    {
                        var desktopStructurePath = System.IO.Path.Combine(desktopFolderPath, "02_Page_Structure_ElementList.txt");
                        await System.IO.File.WriteAllTextAsync(desktopStructurePath, pageStructureReport, System.Text.Encoding.UTF8);
                    }
                    LogToConsole($"  -> Sayfa yapisi raporu kaydedildi");
                    savedFiles++;
                }
                SetProgressValue(50, "Sayfa yapısı tamamlandı", "Adım 4/8 tamamlandı");

                // Step 5: Source code (50-60%)
                SetProgressValue(52, "Kaynak kod toplanıyor...", "Adım 5/8");
                LogToConsole("[5/8] Kaynak kod toplanıyor (web sayfasi ise)...");
                var sourceCode = await CaptureSourceCode(wpfPoint);
                if (!string.IsNullOrEmpty(sourceCode))
                {
                    if (saveToArchive)
                    {
                        var archiveSourcePath = System.IO.Path.Combine(archiveFolderPath, "03_SourceCode.html");
                        await System.IO.File.WriteAllTextAsync(archiveSourcePath, sourceCode, System.Text.Encoding.UTF8);
                        savedFilePaths.Add(archiveSourcePath);
                    }
                    if (saveToDesktop)
                    {
                        var desktopSourcePath = System.IO.Path.Combine(desktopFolderPath, "03_SourceCode.html");
                        await System.IO.File.WriteAllTextAsync(desktopSourcePath, sourceCode, System.Text.Encoding.UTF8);
                    }
                    LogToConsole($"  -> Kaynak kod kaydedildi");
                    savedFiles++;
                }
                else
                {
                    LogToConsole($"  -> Kaynak kod alinamadi (web sayfasi degil veya erisim yok)");
                }
                SetProgressValue(60, "Kaynak kod işlendi", "Adım 5/8 tamamlandı");

                // Step 6: Full screen screenshot (60-75%)
                SetProgressValue(65, "Tam ekran görüntüsü alınıyor...", "Adım 6/8");
                LogToConsole("[6/8] Tum ekran goruntusu aliniyor...");
                if (saveToArchive)
                {
                    var archiveScreenshotPath = System.IO.Path.Combine(archiveFolderPath, "04_Screenshot_FullScreen.png");
                    await CaptureFullScreenToFile(archiveScreenshotPath);
                    savedFilePaths.Add(archiveScreenshotPath);
                }
                if (saveToDesktop)
                {
                    var desktopScreenshotPath = System.IO.Path.Combine(desktopFolderPath, "04_Screenshot_FullScreen.png");
                    await CaptureFullScreenToFile(desktopScreenshotPath);
                }
                LogToConsole($"  -> Tum ekran goruntusu kaydedildi");
                savedFiles++;
                SetProgressValue(75, "Tam ekran görüntüsü alındı", "Adım 6/8 tamamlandı");

                // Step 7: Window screenshot (75-85%)
                SetProgressValue(78, "Pencere görüntüsü alınıyor...", "Adım 7/8");
                LogToConsole("[7/8] Pencere goruntusu aliniyor...");
                if (saveToArchive)
                {
                    var archiveWindowPath = System.IO.Path.Combine(archiveFolderPath, "05_Screenshot_Window.png");
                    await CaptureWindowAtPointToFile(mousePos, archiveWindowPath);
                    savedFilePaths.Add(archiveWindowPath);
                }
                if (saveToDesktop)
                {
                    var desktopWindowPath = System.IO.Path.Combine(desktopFolderPath, "05_Screenshot_Window.png");
                    await CaptureWindowAtPointToFile(mousePos, desktopWindowPath);
                }
                LogToConsole($"  -> Pencere goruntusu kaydedildi");
                savedFiles++;
                SetProgressValue(85, "Pencere görüntüsü alındı", "Adım 7/8 tamamlandı");

                // Step 8: Element screenshot (85-95%)
                SetProgressValue(88, "Element görüntüsü alınıyor...", "Adım 8/8");
                LogToConsole("[8/8] Element goruntusu aliniyor...");
                if (_currentElement != null && _currentElement.Width > 0 && _currentElement.Height > 0)
                {
                    var elementRect = new System.Windows.Rect(_currentElement.X, _currentElement.Y, _currentElement.Width, _currentElement.Height);
                    if (saveToArchive)
                    {
                        var archiveElementPath = System.IO.Path.Combine(archiveFolderPath, "06_Screenshot_Element.png");
                        await CaptureElementToFile(elementRect, archiveElementPath);
                        savedFilePaths.Add(archiveElementPath);
                    }
                    if (saveToDesktop)
                    {
                        var desktopElementPath = System.IO.Path.Combine(desktopFolderPath, "06_Screenshot_Element.png");
                        await CaptureElementToFile(elementRect, desktopElementPath);
                    }
                    LogToConsole($"  -> Element goruntusu kaydedildi");
                    savedFiles++;
                }
                else
                {
                    LogToConsole($"  -> Element goruntusu alinamadi (element bilgisi yok veya boyut gecersiz)");
                }
                SetProgressValue(95, "Element görüntüsü alındı", "Adım 8/8 tamamlandı");

                // Final step: Summary file (95-100%)
                SetProgressValue(97, "Özet dosyası oluşturuluyor...", "Son adım");
                var summaryContent = GenerateSummaryContent(mousePos, savedFiles);
                if (saveToArchive)
                {
                    var archiveSummaryPath = System.IO.Path.Combine(archiveFolderPath, "00_SUMMARY.txt");
                    await System.IO.File.WriteAllTextAsync(archiveSummaryPath, summaryContent, System.Text.Encoding.UTF8);
                    savedFilePaths.Add(archiveSummaryPath);
                }
                if (saveToDesktop)
                {
                    var desktopSummaryPath = System.IO.Path.Combine(desktopFolderPath, "00_SUMMARY.txt");
                    await System.IO.File.WriteAllTextAsync(desktopSummaryPath, summaryContent, System.Text.Encoding.UTF8);
                }

                // Update archive item
                if (saveToArchive && archiveItem != null)
                {
                    archiveItem.FilePaths = savedFilePaths;
                    archiveItem.FileCount = savedFilePaths.Count;
                    _archiveManager.SaveIndex();

                    // Copy archive folder path to clipboard (must be on UI thread)
                    Dispatcher.Invoke(() =>
                    {
                        try
                        {
                            System.Windows.Clipboard.SetText(archiveFolderPath);
                            LogToConsole($"Arsiv klasor linki panoya kopyalandi!");
                        }
                        catch { }
                    });

                    Dispatcher.Invoke(() => RefreshArchiveList());
                }

                SetProgressValue(100, "TAM YAKALAMA TAMAMLANDI!", "Başarılı");
                await Task.Delay(500); // Brief pause to show 100%
                ClearOperationStatus();

                LogToConsole("===========================================");
                LogToConsole($"TAM YAKALAMA TAMAMLANDI: {savedFiles} dosya");
                if (saveToArchive) LogToConsole($"  Arsiv: {archiveFolderPath}");
                if (saveToDesktop) LogToConsole($"  Masaustu: {desktopFolderPath}");
                LogToConsole("===========================================");

                var message = $"Tam yakalama tamamlandi!\n\n" +
                    $"Dosya sayisi: {savedFiles}\n";
                if (saveToArchive) message += $"Arsiv: {archiveFolderPath}\n";
                if (saveToDesktop) message += $"Masaustu: {desktopFolderPath}\n";
                if (saveToArchive) message += $"\nArsiv linki panoya kopyalandi!";

                System.Windows.MessageBox.Show(message, "Tam Yakalama Basarili", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                ClearOperationStatus();
                LogToConsole($"Tam yakalama hatasi: {ex.Message}", Core.Utils.LogLevel.Error);
                System.Windows.MessageBox.Show($"Tam yakalama hatasi:\n{ex.Message}", "Hata", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Capture full screen to file
        /// </summary>
        private async Task CaptureFullScreenToFile(string filePath)
        {
            await Task.Run(() =>
            {
                try
                {
                    var screenBounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                    using (var bitmap = new System.Drawing.Bitmap(screenBounds.Width, screenBounds.Height))
                    {
                        using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                        {
                            graphics.CopyFromScreen(0, 0, 0, 0, screenBounds.Size);
                        }
                        bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Full screen capture error: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// Capture window at point to file
        /// </summary>
        private async Task CaptureWindowAtPointToFile(System.Drawing.Point point, string filePath)
        {
            await Task.Run(() =>
            {
                try
                {
                    // Get window handle at point
                    var hwnd = WindowFromPoint(point);
                    if (hwnd == IntPtr.Zero)
                    {
                        // Fallback to full screen if no window found
                        var screenBounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                        using (var bitmap = new System.Drawing.Bitmap(screenBounds.Width, screenBounds.Height))
                        {
                            using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                            {
                                graphics.CopyFromScreen(0, 0, 0, 0, screenBounds.Size);
                            }
                            bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                        return;
                    }

                    // Get root window (top-level window)
                    var rootHwnd = GetAncestor(hwnd, 2); // GA_ROOT = 2
                    if (rootHwnd == IntPtr.Zero) rootHwnd = hwnd;

                    // Get window rect
                    RECT rect;
                    if (GetWindowRect(rootHwnd, out rect))
                    {
                        int width = rect.Right - rect.Left;
                        int height = rect.Bottom - rect.Top;

                        if (width > 0 && height > 0)
                        {
                            using (var bitmap = new System.Drawing.Bitmap(width, height))
                            {
                                using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                                {
                                    graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, new System.Drawing.Size(width, height));
                                }
                                bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Window capture error: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// Capture element region to file
        /// </summary>
        private async Task CaptureElementToFile(System.Windows.Rect boundingRect, string filePath)
        {
            await Task.Run(() =>
            {
                try
                {
                    int x = (int)boundingRect.X;
                    int y = (int)boundingRect.Y;
                    int width = (int)boundingRect.Width;
                    int height = (int)boundingRect.Height;

                    // Ensure minimum size
                    if (width < 1) width = 1;
                    if (height < 1) height = 1;

                    // Ensure within screen bounds
                    var screenBounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                    if (x < 0) { width += x; x = 0; }
                    if (y < 0) { height += y; y = 0; }
                    if (x + width > screenBounds.Width) width = screenBounds.Width - x;
                    if (y + height > screenBounds.Height) height = screenBounds.Height - y;

                    if (width > 0 && height > 0)
                    {
                        using (var bitmap = new System.Drawing.Bitmap(width, height))
                        {
                            using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                            {
                                graphics.CopyFromScreen(x, y, 0, 0, new System.Drawing.Size(width, height));
                            }
                            bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Element capture error: {ex.Message}");
                }
            });
        }

        /// <summary>
        /// Generate summary content for full capture
        /// </summary>
        private string GenerateSummaryContent(System.Drawing.Point mousePos, int savedFiles)
        {
            var summarySb = new System.Text.StringBuilder();
            summarySb.AppendLine("================================================================================");
            summarySb.AppendLine("              UI ELEMENT INSPECTOR - TAM YAKALAMA OZETI");
            summarySb.AppendLine($"              Tarih: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            summarySb.AppendLine("================================================================================");
            summarySb.AppendLine();
            summarySb.AppendLine($"Mouse Pozisyonu: X={mousePos.X}, Y={mousePos.Y}");
            summarySb.AppendLine($"Kaydedilen Dosya Sayisi: {savedFiles}");
            summarySb.AppendLine();
            summarySb.AppendLine("DOSYALAR:");
            summarySb.AppendLine("  01_Element_5Tech_Report.txt  - 5 Teknoloji ile element ozellikleri");
            summarySb.AppendLine("  02_Page_Structure_ElementList.txt - Sayfa yapisi ve element listesi");
            summarySb.AppendLine("  03_SourceCode.html           - Web sayfasi kaynak kodu (varsa)");
            summarySb.AppendLine("  04_Screenshot_FullScreen.png - Tum ekran goruntusu");
            summarySb.AppendLine("  05_Screenshot_Window.png     - Elementin bulundugu pencere goruntusu");
            summarySb.AppendLine("  06_Screenshot_Element.png    - Secilen elementin goruntusu (varsa)");
            summarySb.AppendLine();
            summarySb.AppendLine("KISAYOL TUSLARI:");
            summarySb.AppendLine("  F1  = Start Inspection (Pencere Gizlenir)");
            summarySb.AppendLine("  F2  = Stop Inspection");
            summarySb.AppendLine("  F3  = Start Inspection (Pencere Gorunur)");
            summarySb.AppendLine("  F4  = DEKLANSOR (Basili Tut = Aktif)");
            summarySb.AppendLine("  F5  = Refresh Element");
            summarySb.AppendLine("  F6  = Masaustu + Arsiv (TXT Rapor)");
            summarySb.AppendLine("  F7  = TAM YAKALAMA (Masaustu + Arsiv)");
            summarySb.AppendLine("  F8  = SADECE ARSIV (Tam Yakalama)");
            summarySb.AppendLine("  Ctrl+S = Hizli Export");
            return summarySb.ToString();
        }

        /// <summary>
        /// Keep the old method name for backward compatibility
        /// </summary>
        private void ExportToDesktop_Click(object sender, RoutedEventArgs e)
        {
            ExportToDesktopAndArchive_Click(sender, e);
        }

        /// <summary>
        /// 5 Teknoloji ile element bilgilerini topla
        /// </summary>
        private async Task<string> CaptureAllTechnologiesReport(System.Windows.Point point)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("================================================================================");
            sb.AppendLine("           ELEMENT RAPORU - 5 TEKNOLOJI ILE TAM ANALIZ");
            sb.AppendLine($"           Tarih: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"           Pozisyon: X={point.X}, Y={point.Y}");
            sb.AppendLine("================================================================================");
            sb.AppendLine();

            var profile = GetSelectedProfile();
            var allResults = new List<ElementInfo>();

            // Her detector ile dene
            foreach (var detector in _detectors)
            {
                try
                {
                    sb.AppendLine($"### {detector.Name.ToUpper()} ###");
                    sb.AppendLine(new string('-', 60));

                    // Playwright icin ozel kontrol - Auto mode'da baslatma
                    if (detector.Name == "Playwright")
                    {
                        sb.AppendLine("  [Playwright atlandi - manuel mod gerektirir]");
                        sb.AppendLine();
                        continue;
                    }

                    if (!detector.CanDetect(point))
                    {
                        sb.AppendLine("  [Bu noktada tespit yapilamadi]");
                        sb.AppendLine();
                        continue;
                    }

                    var element = await detector.GetElementAtPoint(point, profile);
                    if (element == null)
                    {
                        sb.AppendLine("  [Element alinamadi]");
                        sb.AppendLine();
                        continue;
                    }

                    allResults.Add(element);

                    // Temel Bilgiler
                    sb.AppendLine($"  Name: {element.Name}");
                    sb.AppendLine($"  AutomationId: {element.AutomationId}");
                    sb.AppendLine($"  ClassName: {element.ClassName}");
                    sb.AppendLine($"  ControlType: {element.ControlType}");
                    sb.AppendLine($"  LocalizedControlType: {element.LocalizedControlType}");
                    sb.AppendLine($"  ElementType: {element.ElementType}");
                    sb.AppendLine();

                    // Konum
                    sb.AppendLine($"  Konum: X={element.X}, Y={element.Y}");
                    sb.AppendLine($"  Boyut: {element.Width}x{element.Height}");
                    sb.AppendLine($"  BoundingRect: {element.BoundingRectangle}");
                    sb.AppendLine();

                    // Durum
                    sb.AppendLine($"  IsVisible: {element.IsVisible}");
                    sb.AppendLine($"  IsEnabled: {element.IsEnabled}");
                    sb.AppendLine($"  IsOffscreen: {element.IsOffscreen}");
                    sb.AppendLine($"  HasKeyboardFocus: {element.HasKeyboardFocus}");
                    sb.AppendLine();

                    // Handle
                    if (element.WindowHandle != IntPtr.Zero)
                    {
                        sb.AppendLine($"  WindowHandle: 0x{element.WindowHandle.ToInt64():X}");
                        sb.AppendLine($"  WindowTitle: {element.WindowTitle}");
                        sb.AppendLine($"  WindowClassName: {element.WindowClassName}");
                        sb.AppendLine($"  ProcessId: {element.ProcessId}");
                        sb.AppendLine();
                    }

                    // Web/HTML
                    if (!string.IsNullOrEmpty(element.TagName))
                    {
                        sb.AppendLine("  [WEB/HTML]");
                        sb.AppendLine($"    TagName: {element.TagName}");
                        sb.AppendLine($"    HTML Id: {element.HtmlId}");
                        sb.AppendLine($"    HTML Class: {element.HtmlClassName}");
                        sb.AppendLine($"    Href: {element.Href}");
                        sb.AppendLine($"    DocumentUrl: {element.DocumentUrl}");
                        if (!string.IsNullOrEmpty(element.InnerText))
                        {
                            var text = element.InnerText.Length > 200 ? element.InnerText.Substring(0, 200) + "..." : element.InnerText;
                            sb.AppendLine($"    InnerText: {text}");
                        }
                        sb.AppendLine();
                    }

                    // Selectors
                    sb.AppendLine("  [SELECTORS]");
                    if (!string.IsNullOrEmpty(element.XPath)) sb.AppendLine($"    XPath: {element.XPath}");
                    if (!string.IsNullOrEmpty(element.CssSelector)) sb.AppendLine($"    CSS: {element.CssSelector}");
                    if (!string.IsNullOrEmpty(element.PlaywrightSelector)) sb.AppendLine($"    Playwright: {element.PlaywrightSelector}");
                    sb.AppendLine();

                    // Tablo bilgileri
                    if (element.RowIndex >= 0 || element.ColumnIndex >= 0)
                    {
                        sb.AppendLine("  [TABLO BILGILERI]");
                        sb.AppendLine($"    RowIndex: {element.RowIndex}");
                        sb.AppendLine($"    ColumnIndex: {element.ColumnIndex}");
                        sb.AppendLine($"    RowCount: {element.RowCount}");
                        sb.AppendLine($"    ColumnCount: {element.ColumnCount}");
                        sb.AppendLine();
                    }

                    sb.AppendLine();
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"  [HATA: {ex.Message}]");
                    sb.AppendLine();
                }
            }

            // Ozet
            sb.AppendLine("================================================================================");
            sb.AppendLine($"OZET: {allResults.Count} teknolojiden veri toplandi");
            sb.AppendLine("================================================================================");

            return sb.ToString();
        }

        /// <summary>
        /// Sayfa yapisi ve tum elementlerin listesini topla
        /// </summary>
        private async Task<string> CapturePageStructureReport(System.Windows.Point point)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("================================================================================");
            sb.AppendLine("           SAYFA YAPISI VE ELEMENT LISTESI");
            sb.AppendLine($"           Tarih: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine("================================================================================");
            sb.AppendLine();

            try
            {
                // Oncelikle hangi pencerenin uzerinde oldugunu bul
                var windowHandle = GetWindowAtPoint(point);
                if (windowHandle == IntPtr.Zero)
                {
                    sb.AppendLine("[Pencere bulunamadi]");
                    return sb.ToString();
                }

                sb.AppendLine($"Pencere Handle: 0x{windowHandle.ToInt64():X}");

                // Pencere bilgilerini al
                var windowTitle = GetWindowTitle(windowHandle);
                var windowClass = GetWindowClassName(windowHandle);
                sb.AppendLine($"Pencere Basligi: {windowTitle}");
                sb.AppendLine($"Pencere Sinifi: {windowClass}");
                sb.AppendLine();

                // UI Automation ile tum elementleri listele
                sb.AppendLine("ELEMENT LISTESI (UI Automation)");
                sb.AppendLine(new string('=', 80));
                sb.AppendLine();
                sb.AppendLine(String.Format("{0,-5} {1,-25} {2,-20} {3,-15} {4}",
                    "No", "ControlType", "Name", "AutomationId", "ClassName"));
                sb.AppendLine(new string('-', 100));

                var uiaDetector = _detectors.FirstOrDefault(d => d.Name == "UI Automation");
                if (uiaDetector != null)
                {
                    var profile = CollectionProfile.Quick;
                    var elements = await uiaDetector.GetAllElements(windowHandle, profile);

                    int index = 1;
                    foreach (var elem in elements.Take(500)) // Max 500 element
                    {
                        var name = elem.Name ?? "";
                        if (name.Length > 22) name = name.Substring(0, 22) + "...";

                        var autoId = elem.AutomationId ?? "";
                        if (autoId.Length > 12) autoId = autoId.Substring(0, 12) + "...";

                        var className = elem.ClassName ?? "";
                        if (className.Length > 20) className = className.Substring(0, 20) + "...";

                        sb.AppendLine(String.Format("{0,-5} {1,-25} {2,-20} {3,-15} {4}",
                            index++,
                            elem.ControlType ?? elem.ElementType ?? "Unknown",
                            name,
                            autoId,
                            className));
                    }

                    sb.AppendLine();
                    sb.AppendLine($"Toplam Element: {elements.Count}");
                    if (elements.Count > 500)
                        sb.AppendLine($"(Ilk 500 element gosterildi)");
                }
                else
                {
                    sb.AppendLine("[UI Automation detector bulunamadi]");
                }

                // Element hiyerarsisi (agac yapisi)
                sb.AppendLine();
                sb.AppendLine("ELEMENT HIYERARSISI");
                sb.AppendLine(new string('=', 80));
                await AppendElementHierarchy(sb, windowHandle, 0);
            }
            catch (Exception ex)
            {
                sb.AppendLine($"[Sayfa yapisi alinirken hata: {ex.Message}]");
            }

            return sb.ToString();
        }

        private async Task AppendElementHierarchy(System.Text.StringBuilder sb, IntPtr windowHandle, int depth)
        {
            if (depth > 10) return; // Max derinlik

            try
            {
                var root = System.Windows.Automation.AutomationElement.FromHandle(windowHandle);
                await AppendElementNode(sb, root, depth, 100); // Max 100 element per level
            }
            catch (Exception ex)
            {
                sb.AppendLine($"{"".PadLeft(depth * 2)}[Hiyerarsi alinamadi: {ex.Message}]");
            }
        }

        private async Task AppendElementNode(System.Text.StringBuilder sb, System.Windows.Automation.AutomationElement element, int depth, int maxChildren)
        {
            if (element == null || depth > 10) return;

            var indent = "".PadLeft(depth * 2);
            try
            {
                var name = element.Current.Name ?? "";
                if (name.Length > 30) name = name.Substring(0, 30) + "...";

                var controlType = element.Current.ControlType?.ProgrammaticName?.Replace("ControlType.", "") ?? "Unknown";
                var autoId = element.Current.AutomationId ?? "";

                sb.AppendLine($"{indent}[{controlType}] {name} ({autoId})");

                // Alt elementleri al
                var children = element.FindAll(System.Windows.Automation.TreeScope.Children,
                    System.Windows.Automation.Condition.TrueCondition);

                int count = 0;
                foreach (System.Windows.Automation.AutomationElement child in children)
                {
                    if (count++ >= maxChildren)
                    {
                        sb.AppendLine($"{indent}  ... ve {children.Count - maxChildren} element daha");
                        break;
                    }
                    await AppendElementNode(sb, child, depth + 1, maxChildren / 2);
                }
            }
            catch
            {
                sb.AppendLine($"{indent}[Element alinamadi]");
            }
        }

        /// <summary>
        /// Kaynak kodu al (web sayfasi ise)
        /// </summary>
        private async Task<string> CaptureSourceCode(System.Windows.Point point)
        {
            try
            {
                // Oncelikle current element'ten dene
                if (_currentElement != null && !string.IsNullOrEmpty(_currentElement.SourceCode))
                {
                    return _currentElement.SourceCode;
                }

                // Full profile ile element al ve kaynak kodu kontrol et
                var profile = CollectionProfile.Full;

                // MSHTML detector ile dene
                var mshtmlDetector = _detectors.FirstOrDefault(d => d.Name == "MSHTML");
                if (mshtmlDetector != null && mshtmlDetector.CanDetect(point))
                {
                    var element = await mshtmlDetector.GetElementAtPoint(point, profile);
                    if (element != null && !string.IsNullOrEmpty(element.SourceCode))
                    {
                        return element.SourceCode;
                    }
                }

                // WebView2 detector ile dene
                var webviewDetector = _detectors.FirstOrDefault(d => d.Name == "WebView2/CDP");
                if (webviewDetector != null && webviewDetector.CanDetect(point))
                {
                    var element = await webviewDetector.GetElementAtPoint(point, profile);
                    if (element != null && !string.IsNullOrEmpty(element.SourceCode))
                    {
                        return element.SourceCode;
                    }
                }

                return null; // Web sayfasi degil veya kaynak kod alinamadi
            }
            catch (Exception ex)
            {
                LogToConsole($"Kaynak kod alinirken hata: {ex.Message}", Core.Utils.LogLevel.Warning);
                return null;
            }
        }

        /// <summary>
        /// Ekran goruntusunu dosyaya kaydet
        /// </summary>
        private async Task CaptureScreenshotToFile(string filePath)
        {
            await Task.Run(() =>
            {
                try
                {
                    var screenBounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                    using (var bitmap = new System.Drawing.Bitmap(screenBounds.Width, screenBounds.Height))
                    {
                        using (var graphics = System.Drawing.Graphics.FromImage(bitmap))
                        {
                            graphics.CopyFromScreen(0, 0, 0, 0, screenBounds.Size);
                        }
                        bitmap.Save(filePath, System.Drawing.Imaging.ImageFormat.Png);
                    }
                }
                catch (Exception ex)
                {
                    LogToConsole($"Screenshot hatasi: {ex.Message}", Core.Utils.LogLevel.Error);
                }
            });
        }

        // Win32 helper methods
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr WindowFromPoint(System.Drawing.Point point);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder lpString, int nMaxCount);

        [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Auto)]
        private static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder lpClassName, int nMaxCount);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        [return: System.Runtime.InteropServices.MarshalAs(System.Runtime.InteropServices.UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hwnd, out RECT lpRect);

        [System.Runtime.InteropServices.StructLayout(System.Runtime.InteropServices.LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        private IntPtr GetWindowAtPoint(System.Windows.Point point)
        {
            return WindowFromPoint(new System.Drawing.Point((int)point.X, (int)point.Y));
        }

        private string GetWindowTitle(IntPtr hwnd)
        {
            var sb = new System.Text.StringBuilder(256);
            GetWindowText(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        private string GetWindowClassName(IntPtr hwnd)
        {
            var sb = new System.Text.StringBuilder(256);
            GetClassName(hwnd, sb, sb.Capacity);
            return sb.ToString();
        }

        #endregion

        private void InitializeDetectors()
        {
            try
            {
                _logger.LogSection("DETECTOR INITIALIZATION");

                // Add UI Automation detector (primary for desktop apps)
                _detectors.Add(new UIAutomationDetector());
                _logger.LogInfo("UI Automation detector added");

                // Add Win32 detector for native Windows API properties (SPY++ functionality)
                _detectors.Add(new Win32Detector());
                _logger.LogInfo("Win32 API detector added");

                // Add WebView2/CDP detector for modern web
                _detectors.Add(new WebView2Detector());
                _logger.LogInfo("WebView2/CDP detector added");

                // Add MSHTML detector for IE and embedded browsers
                _detectors.Add(new MSHTMLDetector());
                _logger.LogInfo("MSHTML detector added");

                // Add Playwright detector for browser automation
                // Note: Playwright will initialize lazily only when needed
                _detectors.Add(new PlaywrightDetector());
                _logger.LogInfo("Playwright detector added");

                _logger.LogInfo($"Total detectors initialized: {_detectors.Count}");
                LogToConsole($"Initialized {_detectors.Count} detector(s): UI Automation, Win32, WebView2/CDP, MSHTML, Playwright");
                LogToConsole($"📄 Log file: {_logger.LogFilePath}");
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, "Detector initialization failed");
                LogToConsole($"Error initializing detectors: {ex.Message}", Core.Utils.LogLevel.Error);
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

            _logger.LogSection("INSPECTION STARTED");

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
            var modeName = ((ComboBoxItem)cmbInspectionMode.SelectedItem).Content.ToString();

            _logger.LogInfo($"Inspection mode: {modeName} (Index: {mode})");
            LogToConsole($"Starting inspection in mode: {modeName}");

            // Hide main window and show floating toolbar for hover and click modes
            if (mode == 0 || mode == 1)
            {
                _logger.LogInfo("Hiding main window, showing floating toolbar");
                HideMainWindowShowFloating();
            }

            try
            {
                switch (mode)
                {
                    case 0: // Hover (Real-time)
                        _mouseHook.StartHook();
                        _logger.LogInfo("Mouse hook started for Hover mode");
                        break;

                    case 1: // Click (Snapshot)
                        _mouseHook.StartHook();
                        _logger.LogInfo("Mouse hook started for Click mode");
                        LogToConsole("🖱️ Click on an element to capture it (one click per element).");
                        break;

                    case 2: // Region Select
                        _logger.LogInfo("Starting region selection");
                        await StartRegionSelection();
                        break;

                    case 3: // Full Window
                        _logger.LogInfo("Starting full window capture");
                        await CaptureFullWindow();
                        break;
                }
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, "Error during inspection start");
                LogToConsole($"Error starting inspection: {ex.Message}", Core.Utils.LogLevel.Error);
                StopInspection_Click(null, null);
            }
        }

        private async void StartKeepVisible_Click(object sender, RoutedEventArgs e)
        {
            if (_isInspecting)
            {
                LogToConsole("Inspection is already running.");
                return;
            }

            _logger.LogSection("INSPECTION STARTED (KEEP VISIBLE MODE)");

            _isInspecting = true;
            _inspectionCts = new CancellationTokenSource();

            // Update UI
            btnStartInspection.IsEnabled = false;
            btnStartKeepVisible.IsEnabled = false;
            btnStopInspection.IsEnabled = true;
            txtStatus.Text = "Inspecting (Window Visible)...";
            txtStatus.Foreground = System.Windows.Media.Brushes.Green;
            sbStatus.Text = "Inspection mode active - Window visible";

            // Clear previous data
            _collectedElements.Clear();
            ClearElementDetails();

            // Get inspection mode
            var mode = cmbInspectionMode.SelectedIndex;
            var modeName = ((ComboBoxItem)cmbInspectionMode.SelectedItem).Content.ToString();

            _logger.LogInfo($"Inspection mode (Keep Visible): {modeName} (Index: {mode})");
            LogToConsole($"Starting inspection in mode: {modeName} (Window will stay visible)");

            // DO NOT hide main window - keep it visible!
            // User can see both the target app and the inspector

            try
            {
                switch (mode)
                {
                    case 0: // Hover (Real-time)
                        _mouseHook.StartHook();
                        _logger.LogInfo("Mouse hook started for Hover mode (Keep Visible)");
                        LogToConsole("Move mouse over elements. Window will stay visible.");
                        break;

                    case 1: // Click (Snapshot)
                        _mouseHook.StartHook();
                        _logger.LogInfo("Mouse hook started for Click mode (Keep Visible)");
                        LogToConsole("🖱️ Click on an element to capture it. Window will stay visible.");
                        break;

                    case 2: // Region Select
                        _logger.LogInfo("Starting region selection (Keep Visible)");
                        await StartRegionSelection();
                        break;

                    case 3: // Full Window
                        _logger.LogInfo("Starting full window capture (Keep Visible)");
                        await CaptureFullWindow();
                        break;
                }
            }
            catch (Exception ex)
            {
                _logger.LogException(ex, "Error during inspection start (Keep Visible)");
                LogToConsole($"Error starting inspection: {ex.Message}", Core.Utils.LogLevel.Error);
                StopInspection_Click(null, null);
            }
        }

        private void StopInspection_Click(object sender, RoutedEventArgs e)
        {
            if (!_isInspecting) return;

            _logger.LogSection("INSPECTION STOPPED");

            _isInspecting = false;
            _inspectionCts?.Cancel();

            // Update UI
            btnStartInspection.IsEnabled = true;
            btnStartKeepVisible.IsEnabled = true;
            btnStopInspection.IsEnabled = false;
            txtStatus.Text = "Ready";
            txtStatus.Foreground = System.Windows.Media.Brushes.Green;
            sbStatus.Text = "Ready";

            // Stop mouse hook
            _mouseHook.StopHook();
            _logger.LogInfo("Mouse hook stopped");

            // Show main window and hide floating toolbar
            ShowMainWindow();
            _logger.LogInfo("Main window restored");

            _logger.LogInfo($"Total elements collected: {_collectedElements.Count}");
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
                var techIndex = cmbDetectionTech.SelectedIndex;

                LogToConsole($"📍 Capture at point ({point.X}, {point.Y}) - Tech Index: {techIndex}");

                ElementInfo element = null;
                var stopwatch = Stopwatch.StartNew();

                // Check if "All Technologies" mode is selected (index 5)
                if (techIndex == 5)
                {
                    LogToConsole("🔍 All Technologies mode detected!");
                    // Collect data from ALL detectors and merge
                    element = await CaptureFromAllDetectors(point, profile);
                }
                else
                {
                    LogToConsole($"🔍 Single detector mode - Index: {techIndex}");
                    // Use single detector (original behavior)
                    var detector = GetSelectedDetector();

                    if (detector == null)
                    {
                        LogToConsole("No detector available for the current target.");
                        UpdateFloatingWindow("No detector available", 0);
                        return;
                    }

                    element = await detector.GetElementAtPoint(point, profile);
                }

                stopwatch.Stop();

                if (element != null)
                {
                    _currentElement = element;

                    // Capture screenshot of element area
                    try
                    {
                        if (element.BoundingRectangle.Width > 0 && element.BoundingRectangle.Height > 0)
                        {
                            using (var screenshot = Core.Utils.ScreenshotHelper.CaptureElement(element.BoundingRectangle))
                            {
                                if (screenshot != null)
                                {
                                    using (var ms = new MemoryStream())
                                    {
                                        screenshot.Save(ms, System.Drawing.Imaging.ImageFormat.Png);
                                        element.Screenshot = ms.ToArray();
                                    }
                                    LogToConsole($"📷 Element screenshot captured ({element.BoundingRectangle.Width}x{element.BoundingRectangle.Height})");
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToConsole($"Screenshot capture failed: {ex.Message}", Core.Utils.LogLevel.Warning);
                    }

                    // Set source code from OuterHTML if available (for web elements)
                    if (string.IsNullOrEmpty(element.SourceCode))
                    {
                        if (!string.IsNullOrEmpty(element.OuterHTML))
                        {
                            element.SourceCode = element.OuterHTML;
                        }
                        else if (!string.IsNullOrEmpty(element.InnerHTML))
                        {
                            element.SourceCode = $"<{element.TagName ?? "element"}>{element.InnerHTML}</{element.TagName ?? "element"}>";
                        }
                        else
                        {
                            // Generate a simple representation for non-web elements
                            var sb = new StringBuilder();
                            sb.AppendLine($"// Element: {element.Name ?? "Unknown"}");
                            sb.AppendLine($"// Type: {element.ElementType ?? element.ControlType ?? "Unknown"}");
                            sb.AppendLine($"// ClassName: {element.ClassName ?? "N/A"}");
                            sb.AppendLine($"// AutomationId: {element.AutomationId ?? "N/A"}");
                            sb.AppendLine($"// Bounds: ({element.X}, {element.Y}, {element.Width}, {element.Height})");
                            if (!string.IsNullOrEmpty(element.Value))
                                sb.AppendLine($"// Value: {element.Value}");
                            element.SourceCode = sb.ToString();
                        }
                    }

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

                        // Update floating window
                        UpdateFloatingWindow($"Detected: {element.ElementType}", _collectedElements.Count);
                    });
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Error capturing element: {ex.Message}");
                _logger?.LogException(ex, "CaptureElementAtPoint");
                UpdateFloatingWindow($"Error: {ex.Message}", _collectedElements.Count);
            }
        }

        private async Task<ElementInfo> CaptureFromAllDetectors(System.Windows.Point point, CollectionProfile profile)
        {
            var tasks = new List<Task<ElementInfo>>();
            var detectorNames = new List<string>();

            LogToConsole($"=== ALL TECHNOLOGIES MODE ===");
            LogToConsole($"Testing {_detectors.Count} detectors at point ({point.X}, {point.Y})");

            // Launch all detectors in parallel
            foreach (var detector in _detectors)
            {
                try
                {
                    LogToConsole($"Testing {detector.Name}...");
                    if (detector.CanDetect(point))
                    {
                        LogToConsole($"✓ {detector.Name} can detect - adding to list");
                        detectorNames.Add(detector.Name);
                        tasks.Add(detector.GetElementAtPoint(point, profile));
                    }
                    else
                    {
                        LogToConsole($"✗ {detector.Name} cannot detect at this point");
                    }
                }
                catch (Exception ex)
                {
                    LogToConsole($"✗ {detector.Name} check failed: {ex.Message}", Core.Utils.LogLevel.Error);
                }
            }

            if (tasks.Count == 0)
            {
                LogToConsole("❌ No detectors available for this point.", Core.Utils.LogLevel.Error);
                return null;
            }

            LogToConsole($"✓ Using {tasks.Count} detectors: {string.Join(", ", detectorNames)}");
            _logger?.LogInfo($"Collecting from {tasks.Count} technologies: {string.Join(", ", detectorNames)}");

            // Wait for all detectors to complete
            var results = await Task.WhenAll(tasks);

            // Filter out null results
            var validResults = results.Where(r => r != null).ToList();

            if (validResults.Count == 0)
            {
                LogToConsole("❌ All detectors failed to capture element.", Core.Utils.LogLevel.Error);
                return null;
            }

            LogToConsole($"✓ Successfully collected from {validResults.Count} detectors");

            // Log what each detector found
            for (int i = 0; i < validResults.Count; i++)
            {
                var result = validResults[i];
                LogToConsole($"  [{i+1}] {result.DetectionMethod}: Name='{result.Name}', AutomationId='{result.AutomationId}', ClassName='{result.ClassName}', ControlType='{result.ControlType}'");
            }

            _logger?.LogInfo($"Merging data from {validResults.Count} sources");

            // Merge all results into a single ElementInfo
            var merged = MergeElementInfo(validResults);

            LogToConsole($"✓ MERGED RESULT: Name='{merged.Name}', AutomationId='{merged.AutomationId}', ClassName='{merged.ClassName}', ControlType='{merged.ControlType}'");

            return merged;
        }

        private ElementInfo MergeElementInfo(List<ElementInfo> elements)
        {
            if (elements == null || elements.Count == 0)
                return null;

            if (elements.Count == 1)
                return elements[0];

            // Start with first element as base
            var merged = new ElementInfo
            {
                Id = Guid.NewGuid(),
                CaptureTime = DateTime.Now,
                DetectionMethod = $"Combined ({string.Join(" + ", elements.Select(e => e.DetectionMethod))})",
                CollectionProfile = elements[0].CollectionProfile
            };

            // Merge data from all elements
            foreach (var element in elements)
            {
                // Basic properties - use first non-null/non-empty value
                merged.ElementType = merged.ElementType ?? element.ElementType;
                merged.Name = merged.Name ?? element.Name;
                merged.ClassName = merged.ClassName ?? element.ClassName;
                merged.Value = merged.Value ?? element.Value;
                merged.Description = merged.Description ?? element.Description;

                // UI Automation properties
                merged.AutomationId = merged.AutomationId ?? element.AutomationId;
                merged.ControlType = merged.ControlType ?? element.ControlType;
                merged.LocalizedControlType = merged.LocalizedControlType ?? element.LocalizedControlType;
                merged.FrameworkId = merged.FrameworkId ?? element.FrameworkId;
                merged.RuntimeId = merged.RuntimeId ?? element.RuntimeId;
                merged.ItemStatus = merged.ItemStatus ?? element.ItemStatus;
                merged.HelpText = merged.HelpText ?? element.HelpText;
                merged.AcceleratorKey = merged.AcceleratorKey ?? element.AcceleratorKey;
                merged.AccessKey = merged.AccessKey ?? element.AccessKey;

                // Position and size - prefer non-zero values
                if (element.BoundingRectangle.Width > 0 && element.BoundingRectangle.Height > 0)
                {
                    merged.BoundingRectangle = element.BoundingRectangle;
                    merged.X = element.X;
                    merged.Y = element.Y;
                    merged.Width = element.Width;
                    merged.Height = element.Height;
                }

                // Web/HTML properties
                merged.TagName = merged.TagName ?? element.TagName;
                merged.HtmlId = merged.HtmlId ?? element.HtmlId;
                merged.HtmlClassName = merged.HtmlClassName ?? element.HtmlClassName;
                merged.InnerText = merged.InnerText ?? element.InnerText;
                merged.InnerHTML = merged.InnerHTML ?? element.InnerHTML;
                merged.OuterHTML = merged.OuterHTML ?? element.OuterHTML;
                merged.Href = merged.Href ?? element.Href;
                merged.Src = merged.Src ?? element.Src;
                merged.Alt = merged.Alt ?? element.Alt;
                merged.Title = merged.Title ?? element.Title;
                merged.Role = merged.Role ?? element.Role;
                merged.AriaLabel = merged.AriaLabel ?? element.AriaLabel;
                merged.AriaDescribedBy = merged.AriaDescribedBy ?? element.AriaDescribedBy;
                merged.AriaLabelledBy = merged.AriaLabelledBy ?? element.AriaLabelledBy;

                // Selectors
                merged.XPath = merged.XPath ?? element.XPath;
                merged.CssSelector = merged.CssSelector ?? element.CssSelector;
                merged.FullXPath = merged.FullXPath ?? element.FullXPath;
                merged.WindowsPath = merged.WindowsPath ?? element.WindowsPath;
                merged.AccessiblePath = merged.AccessiblePath ?? element.AccessiblePath;
                merged.TreePath = merged.TreePath ?? element.TreePath;
                merged.ElementPath = merged.ElementPath ?? element.ElementPath;
                merged.PlaywrightSelector = merged.PlaywrightSelector ?? element.PlaywrightSelector;
                merged.PlaywrightTableSelector = merged.PlaywrightTableSelector ?? element.PlaywrightTableSelector;

                // Hierarchy
                merged.ParentName = merged.ParentName ?? element.ParentName;
                merged.ParentId = merged.ParentId ?? element.ParentId;
                merged.ParentClassName = merged.ParentClassName ?? element.ParentClassName;

                // Table/Grid properties - prefer non-negative values
                if (element.RowIndex >= 0) merged.RowIndex = element.RowIndex;
                if (element.ColumnIndex >= 0) merged.ColumnIndex = element.ColumnIndex;
                if (element.RowCount >= 0) merged.RowCount = element.RowCount;
                if (element.ColumnCount >= 0) merged.ColumnCount = element.ColumnCount;
                if (element.RowSpan > 1) merged.RowSpan = element.RowSpan;
                if (element.ColumnSpan > 1) merged.ColumnSpan = element.ColumnSpan;
                merged.IsTableCell = merged.IsTableCell || element.IsTableCell;
                merged.IsTableHeader = merged.IsTableHeader || element.IsTableHeader;
                merged.TableName = merged.TableName ?? element.TableName;

                // Merge header lists
                foreach (var header in element.ColumnHeaders)
                {
                    if (!merged.ColumnHeaders.Contains(header))
                        merged.ColumnHeaders.Add(header);
                }
                foreach (var header in element.RowHeaders)
                {
                    if (!merged.RowHeaders.Contains(header))
                        merged.RowHeaders.Add(header);
                }

                // Window info
                merged.WindowTitle = merged.WindowTitle ?? element.WindowTitle;
                merged.WindowClassName = merged.WindowClassName ?? element.WindowClassName;
                merged.ApplicationName = merged.ApplicationName ?? element.ApplicationName;
                merged.ApplicationPath = merged.ApplicationPath ?? element.ApplicationPath;
                if (element.ProcessId > 0) merged.ProcessId = element.ProcessId;
                if (element.WindowHandle != IntPtr.Zero) merged.WindowHandle = element.WindowHandle;
                if (element.NativeWindowHandle != IntPtr.Zero) merged.NativeWindowHandle = element.NativeWindowHandle;

                // Document info (MSHTML)
                merged.DocumentTitle = merged.DocumentTitle ?? element.DocumentTitle;
                merged.DocumentUrl = merged.DocumentUrl ?? element.DocumentUrl;
                merged.DocumentDomain = merged.DocumentDomain ?? element.DocumentDomain;
                merged.DocumentReadyState = merged.DocumentReadyState ?? element.DocumentReadyState;

                // Playwright specific
                merged.InputType = merged.InputType ?? element.InputType;
                merged.InputValue = merged.InputValue ?? element.InputValue;

                // Boolean properties - use OR logic (true if any is true)
                merged.IsEnabled = merged.IsEnabled || element.IsEnabled;
                merged.IsVisible = merged.IsVisible || element.IsVisible;
                merged.IsOffscreen = merged.IsOffscreen && element.IsOffscreen; // false if any is not offscreen
                merged.HasKeyboardFocus = merged.HasKeyboardFocus || element.HasKeyboardFocus;
                merged.IsKeyboardFocusable = merged.IsKeyboardFocusable || element.IsKeyboardFocusable;
                merged.IsPassword = merged.IsPassword || element.IsPassword;
                merged.IsChecked = merged.IsChecked || element.IsChecked;
                merged.IsDisabled = merged.IsDisabled || element.IsDisabled;
                merged.IsEditable = merged.IsEditable || element.IsEditable;

                // Merge collections
                foreach (var pattern in element.SupportedPatterns)
                {
                    if (!merged.SupportedPatterns.Contains(pattern))
                        merged.SupportedPatterns.Add(pattern);
                }

                foreach (var attr in element.HtmlAttributes)
                {
                    if (!merged.HtmlAttributes.ContainsKey(attr.Key))
                        merged.HtmlAttributes[attr.Key] = attr.Value;
                }

                foreach (var style in element.ComputedStyles)
                {
                    if (!merged.ComputedStyles.ContainsKey(style.Key))
                        merged.ComputedStyles[style.Key] = style.Value;
                }

                foreach (var aria in element.AriaAttributes)
                {
                    if (!merged.AriaAttributes.ContainsKey(aria.Key))
                        merged.AriaAttributes[aria.Key] = aria.Value;
                }

                foreach (var data in element.DataAttributes)
                {
                    if (!merged.DataAttributes.ContainsKey(data.Key))
                        merged.DataAttributes[data.Key] = data.Value;
                }

                foreach (var custom in element.CustomProperties)
                {
                    if (!merged.CustomProperties.ContainsKey(custom.Key))
                        merged.CustomProperties[custom.Key] = custom.Value;
                }

                // === NEW: Merge Win32 API properties ===
                if (element.Win32_HWND != IntPtr.Zero)
                {
                    merged.Win32_HWND = element.Win32_HWND;
                    merged.Win32_ParentHWND = element.Win32_ParentHWND;
                    merged.Win32_ClassName = merged.Win32_ClassName ?? element.Win32_ClassName;
                    merged.Win32_WindowText = merged.Win32_WindowText ?? element.Win32_WindowText;
                    if (element.Win32_ThreadId > 0) merged.Win32_ThreadId = element.Win32_ThreadId;
                    merged.Win32_WindowStyles = element.Win32_WindowStyles;
                    merged.Win32_WindowStyles_Parsed = merged.Win32_WindowStyles_Parsed ?? element.Win32_WindowStyles_Parsed;
                    merged.Win32_ExtendedStyles = element.Win32_ExtendedStyles;
                    merged.Win32_ExtendedStyles_Parsed = merged.Win32_ExtendedStyles_Parsed ?? element.Win32_ExtendedStyles_Parsed;
                    if (element.Win32_WindowRect.Width > 0) merged.Win32_WindowRect = element.Win32_WindowRect;
                    if (element.Win32_ClientRect.Width > 0) merged.Win32_ClientRect = element.Win32_ClientRect;
                    merged.Win32_ControlId = merged.Win32_ControlId ?? element.Win32_ControlId;
                    merged.Win32_IsVisible = element.Win32_IsVisible ?? merged.Win32_IsVisible;
                    merged.Win32_IsEnabled = element.Win32_IsEnabled ?? merged.Win32_IsEnabled;
                    merged.Win32_IsMaximized = element.Win32_IsMaximized ?? merged.Win32_IsMaximized;
                    merged.Win32_IsMinimized = element.Win32_IsMinimized ?? merged.Win32_IsMinimized;
                    merged.Win32_IsUnicode = element.Win32_IsUnicode ?? merged.Win32_IsUnicode;
                }

                // === NEW: Merge UIA Control Pattern properties ===
                merged.ValuePattern_IsReadOnly = element.ValuePattern_IsReadOnly ?? merged.ValuePattern_IsReadOnly;
                merged.ValuePattern_Value = merged.ValuePattern_Value ?? element.ValuePattern_Value;
                merged.RangeValue_Minimum = element.RangeValue_Minimum ?? merged.RangeValue_Minimum;
                merged.RangeValue_Maximum = element.RangeValue_Maximum ?? merged.RangeValue_Maximum;
                merged.RangeValue_Value = element.RangeValue_Value ?? merged.RangeValue_Value;
                merged.Selection_CanSelectMultiple = element.Selection_CanSelectMultiple ?? merged.Selection_CanSelectMultiple;
                merged.Selection_IsSelectionRequired = element.Selection_IsSelectionRequired ?? merged.Selection_IsSelectionRequired;
                merged.SelectionItem_IsSelected = element.SelectionItem_IsSelected ?? merged.SelectionItem_IsSelected;
                merged.Scroll_HorizontalPercent = element.Scroll_HorizontalPercent ?? merged.Scroll_HorizontalPercent;
                merged.Scroll_VerticalPercent = element.Scroll_VerticalPercent ?? merged.Scroll_VerticalPercent;
                merged.ExpandCollapse_State = merged.ExpandCollapse_State ?? element.ExpandCollapse_State;
                merged.Toggle_State = merged.Toggle_State ?? element.Toggle_State;
                merged.Transform_CanMove = element.Transform_CanMove ?? merged.Transform_CanMove;
                merged.Transform_CanResize = element.Transform_CanResize ?? merged.Transform_CanResize;
                merged.Dock_Position = merged.Dock_Position ?? element.Dock_Position;
                merged.Window_CanMaximize = element.Window_CanMaximize ?? merged.Window_CanMaximize;
                merged.Window_CanMinimize = element.Window_CanMinimize ?? merged.Window_CanMinimize;
                merged.Window_IsModal = element.Window_IsModal ?? merged.Window_IsModal;
                merged.TextPattern_DocumentRange = merged.TextPattern_DocumentRange ?? element.TextPattern_DocumentRange;

                // === NEW: Merge LegacyIAccessible properties ===
                merged.LegacyName = merged.LegacyName ?? element.LegacyName;
                merged.LegacyValue = merged.LegacyValue ?? element.LegacyValue;
                merged.LegacyDescription = merged.LegacyDescription ?? element.LegacyDescription;
                merged.LegacyHelp = merged.LegacyHelp ?? element.LegacyHelp;
                merged.LegacyKeyboardShortcut = merged.LegacyKeyboardShortcut ?? element.LegacyKeyboardShortcut;
                merged.LegacyDefaultAction = merged.LegacyDefaultAction ?? element.LegacyDefaultAction;
                if (element.LegacyRole > 0) merged.LegacyRole = element.LegacyRole;
                if (element.LegacyState > 0) merged.LegacyState = element.LegacyState;
                if (element.LegacyChildCount > 0) merged.LegacyChildCount = element.LegacyChildCount;

                // === NEW: Merge WebView2/CDP specific properties ===
                if (element.ClassList.Count > 0 && merged.ClassList.Count == 0)
                    merged.ClassList = element.ClassList;
                merged.HtmlName = merged.HtmlName ?? element.HtmlName;
                merged.Placeholder = merged.Placeholder ?? element.Placeholder;
                merged.TabIndex = element.TabIndex ?? merged.TabIndex;
                merged.AriaRole = merged.AriaRole ?? element.AriaRole;
                merged.AriaDisabled = element.AriaDisabled ?? merged.AriaDisabled;
                merged.AriaHidden = element.AriaHidden ?? merged.AriaHidden;
                merged.AriaExpanded = element.AriaExpanded ?? merged.AriaExpanded;
                merged.AriaSelected = element.AriaSelected ?? merged.AriaSelected;
                merged.AriaChecked = element.AriaChecked ?? merged.AriaChecked;
                merged.AriaRequired = element.AriaRequired ?? merged.AriaRequired;
                merged.AriaHasPopup = merged.AriaHasPopup ?? element.AriaHasPopup;
                merged.AriaLevel = element.AriaLevel ?? merged.AriaLevel;

                // === NEW: Merge Layout/Box Model properties ===
                if (element.ClientWidth > 0) merged.ClientWidth = element.ClientWidth;
                if (element.ClientHeight > 0) merged.ClientHeight = element.ClientHeight;
                if (element.OffsetWidth > 0) merged.OffsetWidth = element.OffsetWidth;
                if (element.OffsetHeight > 0) merged.OffsetHeight = element.OffsetHeight;
                if (element.ScrollWidth > 0) merged.ScrollWidth = element.ScrollWidth;
                if (element.ScrollHeight > 0) merged.ScrollHeight = element.ScrollHeight;
                merged.BoxModel_Margin = merged.BoxModel_Margin ?? element.BoxModel_Margin;
                merged.BoxModel_Padding = merged.BoxModel_Padding ?? element.BoxModel_Padding;
                merged.BoxModel_Border = merged.BoxModel_Border ?? element.BoxModel_Border;

                // === NEW: Merge CSS Style properties ===
                merged.Style_Display = merged.Style_Display ?? element.Style_Display;
                merged.Style_Position = merged.Style_Position ?? element.Style_Position;
                merged.Style_Visibility = merged.Style_Visibility ?? element.Style_Visibility;
                merged.Style_Opacity = merged.Style_Opacity ?? element.Style_Opacity;
                merged.Style_Color = merged.Style_Color ?? element.Style_Color;
                merged.Style_BackgroundColor = merged.Style_BackgroundColor ?? element.Style_BackgroundColor;
                merged.Style_FontSize = merged.Style_FontSize ?? element.Style_FontSize;
                merged.Style_FontWeight = merged.Style_FontWeight ?? element.Style_FontWeight;
                merged.Style_ZIndex = merged.Style_ZIndex ?? element.Style_ZIndex;
                merged.Style_Overflow = merged.Style_Overflow ?? element.Style_Overflow;
                merged.Style_Transform = merged.Style_Transform ?? element.Style_Transform;

                // === NEW: Merge MSHTML specific properties ===
                merged.MSHTML_ParentElementTag = merged.MSHTML_ParentElementTag ?? element.MSHTML_ParentElementTag;
                merged.MSHTML_ChildrenCount = element.MSHTML_ChildrenCount ?? merged.MSHTML_ChildrenCount;
                merged.MSHTML_ClientLeft = element.MSHTML_ClientLeft ?? merged.MSHTML_ClientLeft;
                merged.MSHTML_ClientTop = element.MSHTML_ClientTop ?? merged.MSHTML_ClientTop;
                merged.MSHTML_CurrentStyle = merged.MSHTML_CurrentStyle ?? element.MSHTML_CurrentStyle;
                merged.MSHTML_DefaultValue = merged.MSHTML_DefaultValue ?? element.MSHTML_DefaultValue;
                merged.MSHTML_SelectedIndex = element.MSHTML_SelectedIndex ?? merged.MSHTML_SelectedIndex;
                if (element.MSHTML_Options.Count > 0 && merged.MSHTML_Options.Count == 0)
                    merged.MSHTML_Options = element.MSHTML_Options;
                merged.DocumentCookie = merged.DocumentCookie ?? element.DocumentCookie;
                merged.DocumentFramesCount = element.DocumentFramesCount ?? merged.DocumentFramesCount;
                merged.DocumentScriptsCount = element.DocumentScriptsCount ?? merged.DocumentScriptsCount;
                merged.DocumentLinksCount = element.DocumentLinksCount ?? merged.DocumentLinksCount;
                merged.DocumentImagesCount = element.DocumentImagesCount ?? merged.DocumentImagesCount;
                merged.DocumentFormsCount = element.DocumentFormsCount ?? merged.DocumentFormsCount;
                merged.DocumentActiveElement = merged.DocumentActiveElement ?? element.DocumentActiveElement;

                // === NEW: Additional universal state properties ===
                merged.Role = merged.Role ?? element.Role;
                merged.Tag = merged.Tag ?? element.Tag;
                merged.IsSelected = merged.IsSelected || element.IsSelected;
                merged.IsFocused = merged.IsFocused || element.IsFocused;
                merged.IsExpanded = merged.IsExpanded || element.IsExpanded;
                merged.Culture = merged.Culture ?? element.Culture;
                merged.LabeledBy = merged.LabeledBy ?? element.LabeledBy;
                merged.Orientation = merged.Orientation ?? element.Orientation;
                merged.ItemType = merged.ItemType ?? element.ItemType;

                // === Merge technologies used list ===
                foreach (var tech in element.TechnologiesUsed)
                {
                    if (!merged.TechnologiesUsed.Contains(tech))
                        merged.TechnologiesUsed.Add(tech);
                }

                // Merge errors
                foreach (var error in element.CollectionErrors)
                {
                    merged.CollectionErrors.Add($"[{element.DetectionMethod}] {error}");
                }
            }

            LogToConsole($"Merged element data from {elements.Count} sources using {merged.TechnologiesUsed.Count} technologies");
            _logger?.LogInfo($"Successfully merged element data: {merged.ElementType} - {merged.Name} (Technologies: {string.Join(", ", merged.TechnologiesUsed)})");

            return merged;
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
            return GenerateFullReport(element);
        }

        /// <summary>
        /// Generates a comprehensive report with ALL properties from ALL 5 technologies
        /// </summary>
        private string GenerateFullReport(ElementInfo element)
        {
            var sb = new StringBuilder();

            // Header
            sb.AppendLine("╔════════════════════════════════════════════════════════════════════════════════╗");
            sb.AppendLine("║           UNIVERSAL UI ELEMENT INSPECTOR - FULL PROPERTY REPORT               ║");
            sb.AppendLine("╚════════════════════════════════════════════════════════════════════════════════╝");
            sb.AppendLine();
            sb.AppendLine($"  Report Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"  Capture Time: {element.CaptureTime:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"  Detection Method: {element.DetectionMethod}");
            sb.AppendLine($"  Technologies Used: {(element.TechnologiesUsed?.Count > 0 ? string.Join(", ", element.TechnologiesUsed) : "N/A")}");
            sb.AppendLine($"  Collection Profile: {element.CollectionProfile}");
            sb.AppendLine($"  Collection Duration: {element.CollectionDuration.TotalMilliseconds}ms");
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 1. BASIC / UNIVERSAL PROPERTIES
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 1. BASIC / UNIVERSAL PROPERTIES                                                 │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");
            AppendProperty(sb, "Element Type", element.ElementType);
            AppendProperty(sb, "Name", element.Name);
            AppendProperty(sb, "Class Name", element.ClassName);
            AppendProperty(sb, "Value", element.Value);
            AppendProperty(sb, "Description", element.Description);
            AppendProperty(sb, "Role", element.Role);
            AppendProperty(sb, "Tag", element.Tag);
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 2. UI AUTOMATION (UIA) PROPERTIES
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 2. UI AUTOMATION (UIA) PROPERTIES                                               │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");

            sb.AppendLine("  ── 2.1 Basic UIA Properties ──");
            AppendProperty(sb, "AutomationId", element.AutomationId);
            AppendProperty(sb, "Control Type", element.ControlType);
            AppendProperty(sb, "Localized Control Type", element.LocalizedControlType);
            AppendProperty(sb, "Framework Id", element.FrameworkId);
            AppendPropertyIfNotZero(sb, "Process Id", element.ProcessId);
            AppendProperty(sb, "Runtime Id", element.RuntimeId);
            AppendPropertyIfNotZero(sb, "Native Window Handle", element.NativeWindowHandle, true);
            AppendProperty(sb, "Item Status", element.ItemStatus);
            AppendProperty(sb, "Item Type", element.ItemType);
            AppendProperty(sb, "Help Text", element.HelpText);
            AppendProperty(sb, "Accelerator Key", element.AcceleratorKey);
            AppendProperty(sb, "Access Key", element.AccessKey);
            AppendProperty(sb, "Orientation", element.Orientation);
            AppendProperty(sb, "Culture", element.Culture);
            AppendProperty(sb, "Labeled By", element.LabeledBy);
            sb.AppendLine($"  Is Enabled: {element.IsEnabled}");
            sb.AppendLine($"  Is Offscreen: {element.IsOffscreen}");
            sb.AppendLine($"  Has Keyboard Focus: {element.HasKeyboardFocus}");
            sb.AppendLine($"  Is Keyboard Focusable: {element.IsKeyboardFocusable}");
            sb.AppendLine($"  Is Password: {element.IsPassword}");
            sb.AppendLine($"  Is Required For Form: {element.IsRequiredForForm}");
            sb.AppendLine($"  Is Content Element: {element.IsContentElement}");
            sb.AppendLine($"  Is Control Element: {element.IsControlElement}");
            if (element.SupportedPatterns?.Count > 0)
                sb.AppendLine($"  Supported Patterns: {string.Join(", ", element.SupportedPatterns)}");

            sb.AppendLine();
            sb.AppendLine("  ── 2.2 Control Patterns ──");
            AppendProperty(sb, "ValuePattern.Value", element.ValuePattern_Value);
            AppendPropertyNullable(sb, "ValuePattern.IsReadOnly", element.ValuePattern_IsReadOnly);
            AppendPropertyNullable(sb, "RangeValue.Minimum", element.RangeValue_Minimum);
            AppendPropertyNullable(sb, "RangeValue.Maximum", element.RangeValue_Maximum);
            AppendPropertyNullable(sb, "RangeValue.Value", element.RangeValue_Value);
            AppendPropertyNullable(sb, "RangeValue.SmallChange", element.RangeValue_SmallChange);
            AppendPropertyNullable(sb, "RangeValue.LargeChange", element.RangeValue_LargeChange);
            AppendPropertyNullable(sb, "Selection.CanSelectMultiple", element.Selection_CanSelectMultiple);
            AppendPropertyNullable(sb, "Selection.IsSelectionRequired", element.Selection_IsSelectionRequired);
            if (element.Selection_SelectedItems?.Count > 0)
                sb.AppendLine($"  Selection.SelectedItems: {string.Join(", ", element.Selection_SelectedItems)}");
            AppendPropertyNullable(sb, "SelectionItem.IsSelected", element.SelectionItem_IsSelected);
            AppendProperty(sb, "SelectionItem.Container", element.SelectionItem_Container);
            AppendPropertyNullable(sb, "Scroll.HorizontalPercent", element.Scroll_HorizontalPercent);
            AppendPropertyNullable(sb, "Scroll.VerticalPercent", element.Scroll_VerticalPercent);
            AppendPropertyNullable(sb, "Scroll.HorizontalViewSize", element.Scroll_HorizontalViewSize);
            AppendPropertyNullable(sb, "Scroll.VerticalViewSize", element.Scroll_VerticalViewSize);
            AppendPropertyNullable(sb, "Scroll.HorizontallyScrollable", element.Scroll_HorizontallyScrollable);
            AppendPropertyNullable(sb, "Scroll.VerticallyScrollable", element.Scroll_VerticallyScrollable);
            AppendProperty(sb, "ExpandCollapse.State", element.ExpandCollapse_State);
            AppendProperty(sb, "Toggle.State", element.Toggle_State);
            AppendPropertyNullable(sb, "Transform.CanMove", element.Transform_CanMove);
            AppendPropertyNullable(sb, "Transform.CanResize", element.Transform_CanResize);
            AppendPropertyNullable(sb, "Transform.CanRotate", element.Transform_CanRotate);
            AppendProperty(sb, "Dock.Position", element.Dock_Position);
            AppendPropertyNullable(sb, "Window.CanMaximize", element.Window_CanMaximize);
            AppendPropertyNullable(sb, "Window.CanMinimize", element.Window_CanMinimize);
            AppendPropertyNullable(sb, "Window.IsModal", element.Window_IsModal);
            AppendPropertyNullable(sb, "Window.IsTopmost", element.Window_IsTopmost);
            AppendProperty(sb, "Window.InteractionState", element.Window_InteractionState);
            AppendProperty(sb, "Window.VisualState", element.Window_VisualState);
            AppendProperty(sb, "TextPattern.DocumentRange", element.TextPattern_DocumentRange);
            AppendProperty(sb, "TextPattern.SupportedTextSelection", element.TextPattern_SupportedTextSelection);

            sb.AppendLine();
            sb.AppendLine("  ── 2.3 LegacyIAccessible ──");
            AppendProperty(sb, "Legacy Name", element.LegacyName);
            AppendProperty(sb, "Legacy Value", element.LegacyValue);
            AppendProperty(sb, "Legacy Description", element.LegacyDescription);
            AppendProperty(sb, "Legacy Help", element.LegacyHelp);
            AppendProperty(sb, "Legacy Default Action", element.LegacyDefaultAction);
            AppendProperty(sb, "Legacy Keyboard Shortcut", element.LegacyKeyboardShortcut);
            AppendPropertyIfNotZero(sb, "Legacy Role", element.LegacyRole);
            AppendPropertyIfNotZero(sb, "Legacy State", element.LegacyState);
            AppendPropertyIfNotZero(sb, "Legacy Child Count", element.LegacyChildCount);
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 3. WEBVIEW2 / CDP PROPERTIES
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 3. WEBVIEW2 / CDP (Chrome DevTools Protocol) PROPERTIES                         │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");

            sb.AppendLine("  ── 3.1 HTML DOM Attributes ──");
            AppendProperty(sb, "Tag Name", element.TagName);
            AppendProperty(sb, "HTML Id", element.HtmlId);
            AppendProperty(sb, "HTML Class Name", element.HtmlClassName);
            if (element.ClassList?.Count > 0)
                sb.AppendLine($"  Class List: {string.Join(", ", element.ClassList)}");
            AppendProperty(sb, "HTML Name", element.HtmlName);
            AppendProperty(sb, "Inner Text", TruncateForReport(element.InnerText, 300));
            AppendProperty(sb, "Inner HTML", TruncateForReport(element.InnerHTML, 300));
            AppendProperty(sb, "Href", element.Href);
            AppendProperty(sb, "Src", element.Src);
            AppendProperty(sb, "Alt", element.Alt);
            AppendProperty(sb, "Title", element.Title);
            AppendProperty(sb, "Input Type", element.InputType);
            AppendProperty(sb, "Input Value", element.InputValue);
            AppendProperty(sb, "Placeholder", element.Placeholder);
            AppendPropertyNullable(sb, "Tab Index", element.TabIndex);

            if (element.DataAttributes?.Count > 0)
            {
                sb.AppendLine("  ── Data Attributes (data-*) ──");
                foreach (var attr in element.DataAttributes)
                    sb.AppendLine($"    {attr.Key}: {attr.Value}");
            }

            sb.AppendLine();
            sb.AppendLine("  ── 3.2 ARIA Accessibility Attributes ──");
            AppendProperty(sb, "ARIA Role", element.AriaRole);
            AppendProperty(sb, "ARIA Label", element.AriaLabel);
            AppendProperty(sb, "ARIA Described By", element.AriaDescribedBy);
            AppendProperty(sb, "ARIA Labelled By", element.AriaLabelledBy);
            AppendPropertyNullable(sb, "ARIA Disabled", element.AriaDisabled);
            AppendPropertyNullable(sb, "ARIA Hidden", element.AriaHidden);
            AppendPropertyNullable(sb, "ARIA Expanded", element.AriaExpanded);
            AppendPropertyNullable(sb, "ARIA Selected", element.AriaSelected);
            AppendPropertyNullable(sb, "ARIA Checked", element.AriaChecked);
            AppendPropertyNullable(sb, "ARIA Required", element.AriaRequired);
            AppendProperty(sb, "ARIA Has Popup", element.AriaHasPopup);
            AppendPropertyNullable(sb, "ARIA Level", element.AriaLevel);
            AppendProperty(sb, "ARIA Value Text", element.AriaValueText);

            if (element.AriaAttributes?.Count > 0)
            {
                sb.AppendLine("  ── Other ARIA Attributes ──");
                foreach (var attr in element.AriaAttributes)
                    sb.AppendLine($"    {attr.Key}: {attr.Value}");
            }

            sb.AppendLine();
            sb.AppendLine("  ── 3.3 Layout / Box Model ──");
            AppendPropertyIfNotZero(sb, "Client Width", element.ClientWidth);
            AppendPropertyIfNotZero(sb, "Client Height", element.ClientHeight);
            AppendPropertyIfNotZero(sb, "Offset Width", element.OffsetWidth);
            AppendPropertyIfNotZero(sb, "Offset Height", element.OffsetHeight);
            AppendPropertyIfNotZero(sb, "Scroll Width", element.ScrollWidth);
            AppendPropertyIfNotZero(sb, "Scroll Height", element.ScrollHeight);
            AppendPropertyIfNotZero(sb, "Scroll Top", element.ScrollTop);
            AppendPropertyIfNotZero(sb, "Scroll Left", element.ScrollLeft);
            AppendProperty(sb, "Box Model Margin", element.BoxModel_Margin);
            AppendProperty(sb, "Box Model Padding", element.BoxModel_Padding);
            AppendProperty(sb, "Box Model Border", element.BoxModel_Border);

            sb.AppendLine();
            sb.AppendLine("  ── 3.4 CSS Computed Styles ──");
            AppendProperty(sb, "display", element.Style_Display);
            AppendProperty(sb, "position", element.Style_Position);
            AppendProperty(sb, "visibility", element.Style_Visibility);
            AppendProperty(sb, "opacity", element.Style_Opacity);
            AppendProperty(sb, "color", element.Style_Color);
            AppendProperty(sb, "background-color", element.Style_BackgroundColor);
            AppendProperty(sb, "font-size", element.Style_FontSize);
            AppendProperty(sb, "font-weight", element.Style_FontWeight);
            AppendProperty(sb, "z-index", element.Style_ZIndex);
            AppendProperty(sb, "pointer-events", element.Style_PointerEvents);
            AppendProperty(sb, "overflow", element.Style_Overflow);
            AppendProperty(sb, "transform", element.Style_Transform);

            if (element.ComputedStyles?.Count > 0)
            {
                sb.AppendLine("  ── All Computed Styles ──");
                foreach (var style in element.ComputedStyles.Take(30))
                    sb.AppendLine($"    {style.Key}: {style.Value}");
            }
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 4. MSHTML / IHTMLDocument PROPERTIES
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 4. MSHTML / IHTMLDocument (Trident/IE) PROPERTIES                               │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");

            sb.AppendLine("  ── 4.1 IHTMLElement Properties ──");
            AppendProperty(sb, "Parent Element Tag", element.MSHTML_ParentElementTag);
            AppendPropertyNullable(sb, "Children Count", element.MSHTML_ChildrenCount);
            AppendProperty(sb, "Outer Text", TruncateForReport(element.MSHTML_OuterText, 200));
            AppendProperty(sb, "Language", element.MSHTML_Language);
            AppendProperty(sb, "Lang", element.MSHTML_Lang);
            AppendPropertyNullable(sb, "Source Index", element.MSHTML_SourceIndex);
            AppendProperty(sb, "Scope Name", element.MSHTML_ScopeName);
            AppendPropertyNullable(sb, "Can Have Children", element.MSHTML_CanHaveChildren);
            AppendPropertyNullable(sb, "Can Have HTML", element.MSHTML_CanHaveHTML);
            AppendPropertyNullable(sb, "Is Content Editable", element.MSHTML_IsContentEditable);
            AppendProperty(sb, "Content Editable", element.MSHTML_ContentEditable);
            AppendPropertyNullable(sb, "Tab Index", element.MSHTML_TabIndex);
            AppendProperty(sb, "Access Key", element.MSHTML_AccessKey);

            sb.AppendLine();
            sb.AppendLine("  ── 4.2 IHTMLElement2 Properties ──");
            AppendPropertyNullable(sb, "Client Left", element.MSHTML_ClientLeft);
            AppendPropertyNullable(sb, "Client Top", element.MSHTML_ClientTop);
            AppendProperty(sb, "Current Style", element.MSHTML_CurrentStyle);
            AppendProperty(sb, "Runtime Style", element.MSHTML_RuntimeStyle);
            AppendProperty(sb, "Ready State", element.MSHTML_ReadyState);
            AppendProperty(sb, "Dir (Text Direction)", element.MSHTML_Dir);
            AppendPropertyNullable(sb, "Scroll Left", element.MSHTML_ScrollLeft);
            AppendPropertyNullable(sb, "Scroll Top", element.MSHTML_ScrollTop);
            AppendPropertyNullable(sb, "Scroll Width", element.MSHTML_ScrollWidth);
            AppendPropertyNullable(sb, "Scroll Height", element.MSHTML_ScrollHeight);

            sb.AppendLine();
            sb.AppendLine("  ── 4.3 IHTMLInputElement Properties ──");
            AppendProperty(sb, "Default Value", element.MSHTML_DefaultValue);
            AppendPropertyNullable(sb, "Default Checked", element.MSHTML_DefaultChecked);
            AppendProperty(sb, "Form ID", element.MSHTML_Form_Id);
            AppendProperty(sb, "Form Name", element.MSHTML_Form_Name);
            AppendProperty(sb, "Form Action", element.MSHTML_Form_Action);
            AppendProperty(sb, "Form Method", element.MSHTML_Form_Method);
            AppendPropertyNullable(sb, "Max Length", element.MSHTML_MaxLength);
            AppendPropertyNullable(sb, "Size", element.MSHTML_Size);
            AppendPropertyNullable(sb, "Read Only", element.MSHTML_ReadOnly);
            AppendProperty(sb, "Accept", element.MSHTML_Accept);
            AppendProperty(sb, "Align", element.MSHTML_Align);
            AppendPropertyNullable(sb, "Indeterminate", element.MSHTML_IndeterminateState);

            sb.AppendLine();
            sb.AppendLine("  ── 4.4 IHTMLSelectElement Properties ──");
            AppendPropertyNullable(sb, "Selected Index", element.MSHTML_SelectedIndex);
            AppendProperty(sb, "Selected Value", element.MSHTML_SelectedValue);
            AppendProperty(sb, "Selected Text", element.MSHTML_SelectedText);
            AppendPropertyNullable(sb, "Options Length", element.MSHTML_OptionsLength);
            AppendPropertyNullable(sb, "Multiple", element.MSHTML_Multiple);
            if (element.MSHTML_Options?.Count > 0)
                sb.AppendLine($"  Options: {string.Join(", ", element.MSHTML_Options.Take(20))}");
            if (element.MSHTML_OptionValues?.Count > 0)
                sb.AppendLine($"  Option Values: {string.Join(", ", element.MSHTML_OptionValues.Take(20))}");

            sb.AppendLine();
            sb.AppendLine("  ── 4.5 IHTMLTextAreaElement Properties ──");
            AppendPropertyNullable(sb, "Cols", element.MSHTML_Cols);
            AppendPropertyNullable(sb, "Rows", element.MSHTML_Rows);
            AppendProperty(sb, "Wrap", element.MSHTML_Wrap);

            sb.AppendLine();
            sb.AppendLine("  ── 4.6 IHTMLButtonElement Properties ──");
            AppendProperty(sb, "Button Type", element.MSHTML_ButtonType);
            AppendProperty(sb, "Form Action", element.MSHTML_FormAction);
            AppendProperty(sb, "Form Method", element.MSHTML_FormMethod);

            sb.AppendLine();
            sb.AppendLine("  ── 4.7 IHTMLAnchorElement Properties ──");
            AppendProperty(sb, "Target", element.MSHTML_Target);
            AppendProperty(sb, "Protocol", element.MSHTML_Protocol);
            AppendProperty(sb, "Host", element.MSHTML_Host);
            AppendProperty(sb, "Hostname", element.MSHTML_Hostname);
            AppendProperty(sb, "Port", element.MSHTML_Port);
            AppendProperty(sb, "Pathname", element.MSHTML_Pathname);
            AppendProperty(sb, "Search (Query)", element.MSHTML_Search);
            AppendProperty(sb, "Hash (Fragment)", element.MSHTML_Hash);
            AppendProperty(sb, "Rel", element.MSHTML_Rel);

            sb.AppendLine();
            sb.AppendLine("  ── 4.8 IHTMLImageElement Properties ──");
            AppendPropertyNullable(sb, "Is Map", element.MSHTML_IsMap);
            AppendPropertyNullable(sb, "Natural Width", element.MSHTML_NaturalWidth);
            AppendPropertyNullable(sb, "Natural Height", element.MSHTML_NaturalHeight);
            AppendPropertyNullable(sb, "Complete", element.MSHTML_Complete);
            AppendProperty(sb, "Long Desc", element.MSHTML_LongDesc);

            sb.AppendLine();
            sb.AppendLine("  ── 4.9 IHTMLTableElement Properties ──");
            AppendProperty(sb, "Table Caption", element.MSHTML_Table_Caption);
            AppendProperty(sb, "Table Summary", element.MSHTML_Table_Summary);
            AppendProperty(sb, "Table Border", element.MSHTML_Table_Border);
            AppendProperty(sb, "Cell Padding", element.MSHTML_Table_CellPadding);
            AppendProperty(sb, "Cell Spacing", element.MSHTML_Table_CellSpacing);
            AppendProperty(sb, "Table Width", element.MSHTML_Table_Width);
            AppendProperty(sb, "Table BgColor", element.MSHTML_Table_BgColor);
            AppendPropertyNullable(sb, "Table Rows Count", element.MSHTML_Table_RowsCount);
            AppendPropertyNullable(sb, "Table TBodies Count", element.MSHTML_Table_TBodiesCount);

            sb.AppendLine();
            sb.AppendLine("  ── 4.10 IHTMLTableRowElement Properties ──");
            AppendPropertyNullable(sb, "Row Index", element.MSHTML_Row_RowIndex);
            AppendPropertyNullable(sb, "Section Row Index", element.MSHTML_Row_SectionRowIndex);
            AppendPropertyNullable(sb, "Cells Count", element.MSHTML_Row_CellsCount);
            AppendProperty(sb, "Row BgColor", element.MSHTML_Row_BgColor);
            AppendProperty(sb, "Row VAlign", element.MSHTML_Row_VAlign);
            AppendProperty(sb, "Row Align", element.MSHTML_Row_Align);

            sb.AppendLine();
            sb.AppendLine("  ── 4.11 IHTMLTableCellElement Properties ──");
            AppendPropertyNullable(sb, "Cell Index", element.MSHTML_Cell_CellIndex);
            AppendProperty(sb, "Cell Abbr", element.MSHTML_Cell_Abbr);
            AppendProperty(sb, "Cell Axis", element.MSHTML_Cell_Axis);
            AppendProperty(sb, "Cell Headers", element.MSHTML_Cell_Headers);
            AppendProperty(sb, "Cell Scope", element.MSHTML_Cell_Scope);
            AppendProperty(sb, "Cell NoWrap", element.MSHTML_Cell_NoWrap);
            AppendProperty(sb, "Cell BgColor", element.MSHTML_Cell_BgColor);
            AppendProperty(sb, "Cell VAlign", element.MSHTML_Cell_VAlign);
            AppendProperty(sb, "Cell Align", element.MSHTML_Cell_Align);

            sb.AppendLine();
            sb.AppendLine("  ── 4.12 IHTMLFrameElement Properties ──");
            AppendProperty(sb, "Frame Src", element.MSHTML_Frame_Src);
            AppendProperty(sb, "Frame Name", element.MSHTML_Frame_Name);
            AppendProperty(sb, "Frame Scrolling", element.MSHTML_Frame_Scrolling);
            AppendProperty(sb, "Frame Border", element.MSHTML_Frame_FrameBorder);
            AppendProperty(sb, "Margin Width", element.MSHTML_Frame_MarginWidth);
            AppendProperty(sb, "Margin Height", element.MSHTML_Frame_MarginHeight);
            AppendPropertyNullable(sb, "No Resize", element.MSHTML_Frame_NoResize);

            sb.AppendLine();
            sb.AppendLine("  ── 4.13 IHTMLDocument Properties ──");
            AppendProperty(sb, "Document Title", element.DocumentTitle);
            AppendProperty(sb, "Document URL", element.DocumentUrl);
            AppendProperty(sb, "Document Domain", element.DocumentDomain);
            AppendProperty(sb, "Document Ready State", element.DocumentReadyState);
            AppendProperty(sb, "Document Charset", element.DocumentCharset);
            AppendProperty(sb, "Document Last Modified", element.DocumentLastModified);
            AppendProperty(sb, "Document Referrer", element.DocumentReferrer);
            AppendProperty(sb, "Document Compat Mode", element.DocumentCompatMode);
            AppendProperty(sb, "Document Design Mode", element.DocumentDesignMode);
            AppendProperty(sb, "Document DocType", element.DocumentDocType);
            AppendProperty(sb, "Document Dir", element.DocumentDir);
            AppendProperty(sb, "Document Cookie", TruncateForReport(element.DocumentCookie, 100));
            AppendPropertyNullable(sb, "Frames Count", element.DocumentFramesCount);
            AppendPropertyNullable(sb, "Scripts Count", element.DocumentScriptsCount);
            AppendPropertyNullable(sb, "Links Count", element.DocumentLinksCount);
            AppendPropertyNullable(sb, "Images Count", element.DocumentImagesCount);
            AppendPropertyNullable(sb, "Forms Count", element.DocumentFormsCount);
            AppendProperty(sb, "Active Element", element.DocumentActiveElement);

            sb.AppendLine();
            sb.AppendLine("  ── 4.14 IHTMLDocument2/3/4/5 Properties ──");
            AppendProperty(sb, "Protocol", element.MSHTML_Doc_Protocol);
            AppendProperty(sb, "Name Prop", element.MSHTML_Doc_NameProp);
            AppendProperty(sb, "File Created Date", element.MSHTML_Doc_FileCreatedDate);
            AppendProperty(sb, "File Modified Date", element.MSHTML_Doc_FileModifiedDate);
            AppendProperty(sb, "File Size", element.MSHTML_Doc_FileSize);
            AppendProperty(sb, "Mime Type", element.MSHTML_Doc_MimeType);
            AppendProperty(sb, "Security", element.MSHTML_Doc_Security);
            AppendPropertyNullable(sb, "Anchors Count", element.MSHTML_Doc_Anchors_Count);
            AppendPropertyNullable(sb, "Applets Count", element.MSHTML_Doc_Applets_Count);
            AppendPropertyNullable(sb, "Embeds Count", element.MSHTML_Doc_Embeds_Count);
            AppendPropertyNullable(sb, "All Elements Count", element.MSHTML_Doc_All_Count);
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 5. WIN32 API / SPY++ PROPERTIES
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 5. WIN32 API / SPY++ PROPERTIES                                                 │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");

            sb.AppendLine("  ── 5.1 Window Handle Information ──");
            AppendPropertyIfNotZero(sb, "HWND", element.Win32_HWND, true);
            AppendPropertyIfNotZero(sb, "Parent HWND", element.Win32_ParentHWND, true);
            AppendProperty(sb, "Win32 Class Name", element.Win32_ClassName);
            AppendProperty(sb, "Win32 Window Text", element.Win32_WindowText);
            AppendPropertyIfNotZero(sb, "Thread Id", element.Win32_ThreadId);
            AppendProperty(sb, "Control Id", element.Win32_ControlId);
            AppendPropertyNullable(sb, "Is Unicode", element.Win32_IsUnicode);
            AppendPropertyNullable(sb, "Is Visible", element.Win32_IsVisible);
            AppendPropertyNullable(sb, "Is Enabled", element.Win32_IsEnabled);
            AppendPropertyNullable(sb, "Is Maximized", element.Win32_IsMaximized);
            AppendPropertyNullable(sb, "Is Minimized", element.Win32_IsMinimized);
            AppendPropertyIfNotZero(sb, "Instance Handle", element.Win32_InstanceHandle, true);
            AppendPropertyIfNotZero(sb, "Menu Handle", element.Win32_MenuHandle, true);
            AppendPropertyIfNotZero(sb, "WndProc", element.Win32_WndProc, true);

            sb.AppendLine();
            sb.AppendLine("  ── 5.2 Window Styles (WS_* flags) ──");
            if (element.Win32_WindowStyles > 0)
                sb.AppendLine($"  Window Styles (Raw): 0x{element.Win32_WindowStyles:X8}");
            AppendProperty(sb, "Window Styles (Parsed)", element.Win32_WindowStyles_Parsed);

            sb.AppendLine();
            sb.AppendLine("  ── 5.3 Extended Styles (WS_EX_* flags) ──");
            if (element.Win32_ExtendedStyles > 0)
                sb.AppendLine($"  Extended Styles (Raw): 0x{element.Win32_ExtendedStyles:X8}");
            AppendProperty(sb, "Extended Styles (Parsed)", element.Win32_ExtendedStyles_Parsed);

            sb.AppendLine();
            sb.AppendLine("  ── 5.4 Window Rectangles ──");
            if (element.Win32_WindowRect.Width > 0)
                sb.AppendLine($"  Window Rect: X={element.Win32_WindowRect.X}, Y={element.Win32_WindowRect.Y}, Width={element.Win32_WindowRect.Width}, Height={element.Win32_WindowRect.Height}");
            if (element.Win32_ClientRect.Width > 0)
                sb.AppendLine($"  Client Rect: Width={element.Win32_ClientRect.Width}, Height={element.Win32_ClientRect.Height}");
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 6. POSITION & SIZE
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 6. POSITION & SIZE                                                              │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");
            if (element.BoundingRectangle.Width > 0)
                sb.AppendLine($"  Bounding Rectangle: {element.BoundingRectangle}");
            sb.AppendLine($"  X: {element.X}");
            sb.AppendLine($"  Y: {element.Y}");
            sb.AppendLine($"  Width: {element.Width}");
            sb.AppendLine($"  Height: {element.Height}");
            if (element.ClickablePoint.X > 0 || element.ClickablePoint.Y > 0)
                sb.AppendLine($"  Clickable Point: {element.ClickablePoint.X}, {element.ClickablePoint.Y}");
            if (element.ClientRect.Width > 0)
                sb.AppendLine($"  Client Rect: {element.ClientRect}");
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 7. SELECTORS
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 7. SELECTORS                                                                    │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");
            AppendProperty(sb, "XPath", element.XPath);
            AppendProperty(sb, "Full XPath", element.FullXPath);
            AppendProperty(sb, "CSS Selector", element.CssSelector);
            AppendProperty(sb, "Windows Path", element.WindowsPath);
            AppendProperty(sb, "Accessible Path", element.AccessiblePath);
            AppendProperty(sb, "Tree Path", element.TreePath);
            AppendProperty(sb, "Element Path", element.ElementPath);
            AppendProperty(sb, "Playwright Selector", element.PlaywrightSelector);
            AppendProperty(sb, "Playwright Table Selector", element.PlaywrightTableSelector);
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 8. HIERARCHY
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 8. HIERARCHY                                                                    │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");
            AppendProperty(sb, "Parent Name", element.ParentName);
            AppendProperty(sb, "Parent Id", element.ParentId);
            AppendProperty(sb, "Parent Class Name", element.ParentClassName);
            sb.AppendLine($"  Children Count: {element.Children?.Count ?? 0}");
            sb.AppendLine($"  Child Index: {element.ChildIndex}");
            sb.AppendLine($"  Tree Level: {element.TreeLevel}");
            AppendProperty(sb, "Owner Window", element.OwnerWindow);
            AppendProperty(sb, "Control Container", element.ControlContainer);
            AppendProperty(sb, "Window Title", element.WindowTitle);
            AppendProperty(sb, "Window Class Name", element.WindowClassName);
            AppendProperty(sb, "Application Name", element.ApplicationName);
            AppendProperty(sb, "Application Path", element.ApplicationPath);
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 9. TABLE/GRID PROPERTIES
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 9. TABLE/GRID PROPERTIES                                                        │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");
            sb.AppendLine($"  Is Table Cell: {element.IsTableCell}");
            sb.AppendLine($"  Is Table Header: {element.IsTableHeader}");
            if (element.RowIndex >= 0) sb.AppendLine($"  Row Index: {element.RowIndex}");
            if (element.ColumnIndex >= 0) sb.AppendLine($"  Column Index: {element.ColumnIndex}");
            if (element.RowCount >= 0) sb.AppendLine($"  Row Count: {element.RowCount}");
            if (element.ColumnCount >= 0) sb.AppendLine($"  Column Count: {element.ColumnCount}");
            if (element.RowSpan > 1) sb.AppendLine($"  Row Span: {element.RowSpan}");
            if (element.ColumnSpan > 1) sb.AppendLine($"  Column Span: {element.ColumnSpan}");
            AppendProperty(sb, "Table Name", element.TableName);
            if (element.ColumnHeaders?.Count > 0)
                sb.AppendLine($"  Column Headers: {string.Join(", ", element.ColumnHeaders)}");
            if (element.RowHeaders?.Count > 0)
                sb.AppendLine($"  Row Headers: {string.Join(", ", element.RowHeaders)}");
            AppendProperty(sb, "Row Or Column Major", element.Table_RowOrColumnMajor);
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 10. STATE PROPERTIES
            // ═══════════════════════════════════════════════════════════════════════════════
            sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
            sb.AppendLine("│ 10. STATE PROPERTIES                                                            │");
            sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");
            sb.AppendLine($"  Is Visible: {element.IsVisible}");
            sb.AppendLine($"  Is Hidden: {element.IsHidden}");
            sb.AppendLine($"  Is Checked: {element.IsChecked}");
            sb.AppendLine($"  Is Disabled: {element.IsDisabled}");
            sb.AppendLine($"  Is Editable: {element.IsEditable}");
            sb.AppendLine($"  Is Selected: {element.IsSelected}");
            sb.AppendLine($"  Is Focused: {element.IsFocused}");
            sb.AppendLine($"  Is Expanded: {element.IsExpanded}");
            sb.AppendLine();

            // ═══════════════════════════════════════════════════════════════════════════════
            // 11. CUSTOM / EXTRA PROPERTIES
            // ═══════════════════════════════════════════════════════════════════════════════
            if (element.CustomProperties?.Count > 0)
            {
                sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
                sb.AppendLine("│ 11. CUSTOM / EXTRA PROPERTIES                                                   │");
                sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");
                foreach (var prop in element.CustomProperties)
                {
                    sb.AppendLine($"  {prop.Key}: {prop.Value}");
                }
                sb.AppendLine();
            }

            // ═══════════════════════════════════════════════════════════════════════════════
            // COLLECTION NOTES / ERRORS
            // ═══════════════════════════════════════════════════════════════════════════════
            if (element.CollectionErrors?.Count > 0)
            {
                sb.AppendLine("┌──────────────────────────────────────────────────────────────────────────────────┐");
                sb.AppendLine("│ COLLECTION NOTES / ERRORS                                                       │");
                sb.AppendLine("└──────────────────────────────────────────────────────────────────────────────────┘");
                foreach (var error in element.CollectionErrors)
                {
                    sb.AppendLine($"  ⚠ {error}");
                }
                sb.AppendLine();
            }

            // Footer
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════");
            sb.AppendLine("                              END OF REPORT                                         ");
            sb.AppendLine("════════════════════════════════════════════════════════════════════════════════════");

            return sb.ToString();
        }

        private string TruncateForReport(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return null;
            if (value.Length <= maxLength) return value;
            return value.Substring(0, maxLength) + "...";
        }

        private void AppendPropertyIfNotZero(StringBuilder sb, string name, int value)
        {
            if (value != 0)
                sb.AppendLine($"  {name}: {value}");
        }

        private void AppendPropertyIfNotZero(StringBuilder sb, string name, double value)
        {
            if (value != 0)
                sb.AppendLine($"  {name}: {value}");
        }

        private void AppendPropertyIfNotZero(StringBuilder sb, string name, IntPtr value, bool asHex = false)
        {
            if (value != IntPtr.Zero)
            {
                if (asHex)
                    sb.AppendLine($"  {name}: 0x{value.ToInt64():X8}");
                else
                    sb.AppendLine($"  {name}: {value}");
            }
        }

        private void AppendPropertyNullable<T>(StringBuilder sb, string name, T? value) where T : struct
        {
            if (value.HasValue)
                sb.AppendLine($"  {name}: {value.Value}");
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
            // === 1. BASIC / UNIVERSAL PROPERTIES ===
            var basicProps = new List<KeyValuePair<string, string>>
            {
                new("Element Type", element.ElementType ?? ""),
                new("Name", element.Name ?? ""),
                new("Class Name", element.ClassName ?? ""),
                new("Value", element.Value ?? ""),
                new("Description", element.Description ?? ""),
                new("Role", element.Role ?? ""),
                new("Tag", element.Tag ?? ""),
                new("Detection Method", element.DetectionMethod ?? ""),
                new("Technologies Used", element.TechnologiesUsed?.Count > 0 ? string.Join(", ", element.TechnologiesUsed) : "")
            };
            dgBasicProperties.ItemsSource = basicProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // === 2. UI AUTOMATION (UIA) PROPERTIES ===
            // 2.1 Basic UIA
            var uiaBasicProps = new List<KeyValuePair<string, string>>
            {
                new("AutomationId", element.AutomationId ?? ""),
                new("Control Type", element.ControlType ?? ""),
                new("Localized Control Type", element.LocalizedControlType ?? ""),
                new("Framework Id", element.FrameworkId ?? ""),
                new("Process Id", element.ProcessId > 0 ? element.ProcessId.ToString() : ""),
                new("Runtime Id", element.RuntimeId ?? ""),
                new("Native Window Handle", element.NativeWindowHandle != IntPtr.Zero ? $"0x{element.NativeWindowHandle.ToInt64():X8}" : ""),
                new("Item Status", element.ItemStatus ?? ""),
                new("Item Type", element.ItemType ?? ""),
                new("Help Text", element.HelpText ?? ""),
                new("Accelerator Key", element.AcceleratorKey ?? ""),
                new("Access Key", element.AccessKey ?? ""),
                new("Orientation", element.Orientation ?? ""),
                new("Culture", element.Culture ?? ""),
                new("Labeled By", element.LabeledBy ?? ""),
                new("Is Enabled", element.IsEnabled.ToString()),
                new("Is Offscreen", element.IsOffscreen.ToString()),
                new("Has Keyboard Focus", element.HasKeyboardFocus.ToString()),
                new("Is Keyboard Focusable", element.IsKeyboardFocusable.ToString()),
                new("Is Password", element.IsPassword.ToString()),
                new("Is Required For Form", element.IsRequiredForForm.ToString()),
                new("Is Content Element", element.IsContentElement.ToString()),
                new("Is Control Element", element.IsControlElement.ToString()),
                new("Supported Patterns", element.SupportedPatterns?.Count > 0 ? string.Join(", ", element.SupportedPatterns) : "")
            };
            dgUIABasic.ItemsSource = uiaBasicProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // 2.2 UIA Control Patterns
            var uiaPatternProps = new List<KeyValuePair<string, string>>
            {
                new("ValuePattern.Value", element.ValuePattern_Value ?? ""),
                new("ValuePattern.IsReadOnly", element.ValuePattern_IsReadOnly?.ToString() ?? ""),
                new("RangeValue.Minimum", element.RangeValue_Minimum?.ToString() ?? ""),
                new("RangeValue.Maximum", element.RangeValue_Maximum?.ToString() ?? ""),
                new("RangeValue.Value", element.RangeValue_Value?.ToString() ?? ""),
                new("RangeValue.SmallChange", element.RangeValue_SmallChange?.ToString() ?? ""),
                new("RangeValue.LargeChange", element.RangeValue_LargeChange?.ToString() ?? ""),
                new("Selection.CanSelectMultiple", element.Selection_CanSelectMultiple?.ToString() ?? ""),
                new("Selection.IsSelectionRequired", element.Selection_IsSelectionRequired?.ToString() ?? ""),
                new("Selection.SelectedItems", element.Selection_SelectedItems?.Count > 0 ? string.Join(", ", element.Selection_SelectedItems) : ""),
                new("SelectionItem.IsSelected", element.SelectionItem_IsSelected?.ToString() ?? ""),
                new("SelectionItem.Container", element.SelectionItem_Container ?? ""),
                new("Scroll.HorizontalPercent", element.Scroll_HorizontalPercent?.ToString() ?? ""),
                new("Scroll.VerticalPercent", element.Scroll_VerticalPercent?.ToString() ?? ""),
                new("Scroll.HorizontallyScrollable", element.Scroll_HorizontallyScrollable?.ToString() ?? ""),
                new("Scroll.VerticallyScrollable", element.Scroll_VerticallyScrollable?.ToString() ?? ""),
                new("ExpandCollapse.State", element.ExpandCollapse_State ?? ""),
                new("Toggle.State", element.Toggle_State ?? ""),
                new("Transform.CanMove", element.Transform_CanMove?.ToString() ?? ""),
                new("Transform.CanResize", element.Transform_CanResize?.ToString() ?? ""),
                new("Transform.CanRotate", element.Transform_CanRotate?.ToString() ?? ""),
                new("Dock.Position", element.Dock_Position ?? ""),
                new("Window.CanMaximize", element.Window_CanMaximize?.ToString() ?? ""),
                new("Window.CanMinimize", element.Window_CanMinimize?.ToString() ?? ""),
                new("Window.IsModal", element.Window_IsModal?.ToString() ?? ""),
                new("Window.IsTopmost", element.Window_IsTopmost?.ToString() ?? ""),
                new("Window.InteractionState", element.Window_InteractionState ?? ""),
                new("Window.VisualState", element.Window_VisualState ?? ""),
                new("TextPattern.DocumentRange", element.TextPattern_DocumentRange ?? ""),
                new("TextPattern.SupportedTextSelection", element.TextPattern_SupportedTextSelection ?? "")
            };
            dgUIAPatterns.ItemsSource = uiaPatternProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // 2.3 LegacyIAccessible
            var uiaLegacyProps = new List<KeyValuePair<string, string>>
            {
                new("Legacy Name", element.LegacyName ?? ""),
                new("Legacy Value", element.LegacyValue ?? ""),
                new("Legacy Description", element.LegacyDescription ?? ""),
                new("Legacy Help", element.LegacyHelp ?? ""),
                new("Legacy Default Action", element.LegacyDefaultAction ?? ""),
                new("Legacy Keyboard Shortcut", element.LegacyKeyboardShortcut ?? ""),
                new("Legacy Role", element.LegacyRole > 0 ? element.LegacyRole.ToString() : ""),
                new("Legacy State", element.LegacyState > 0 ? element.LegacyState.ToString() : ""),
                new("Legacy Child Count", element.LegacyChildCount > 0 ? element.LegacyChildCount.ToString() : "")
            };
            dgUIALegacy.ItemsSource = uiaLegacyProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // === 3. WEBVIEW2 / CDP PROPERTIES ===
            // 3.1 HTML DOM Attributes
            var webDOMProps = new List<KeyValuePair<string, string>>
            {
                new("Tag Name", element.TagName ?? ""),
                new("HTML Id", element.HtmlId ?? ""),
                new("HTML Class Name", element.HtmlClassName ?? ""),
                new("Class List", element.ClassList?.Count > 0 ? string.Join(", ", element.ClassList) : ""),
                new("HTML Name", element.HtmlName ?? ""),
                new("Inner Text", TruncateString(element.InnerText, 200)),
                new("Inner HTML", TruncateString(element.InnerHTML, 200)),
                new("Href", element.Href ?? ""),
                new("Src", element.Src ?? ""),
                new("Alt", element.Alt ?? ""),
                new("Title", element.Title ?? ""),
                new("Input Type", element.InputType ?? ""),
                new("Input Value", element.InputValue ?? ""),
                new("Placeholder", element.Placeholder ?? ""),
                new("Tab Index", element.TabIndex?.ToString() ?? "")
            };
            // Add data-* attributes
            if (element.DataAttributes?.Count > 0)
            {
                foreach (var attr in element.DataAttributes.Take(10))
                    webDOMProps.Add(new(attr.Key, attr.Value));
            }
            dgWebDOM.ItemsSource = webDOMProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // 3.2 ARIA Attributes
            var webARIAProps = new List<KeyValuePair<string, string>>
            {
                new("ARIA Role", element.AriaRole ?? ""),
                new("ARIA Label", element.AriaLabel ?? ""),
                new("ARIA Described By", element.AriaDescribedBy ?? ""),
                new("ARIA Labelled By", element.AriaLabelledBy ?? ""),
                new("ARIA Disabled", element.AriaDisabled?.ToString() ?? ""),
                new("ARIA Hidden", element.AriaHidden?.ToString() ?? ""),
                new("ARIA Expanded", element.AriaExpanded?.ToString() ?? ""),
                new("ARIA Selected", element.AriaSelected?.ToString() ?? ""),
                new("ARIA Checked", element.AriaChecked?.ToString() ?? ""),
                new("ARIA Required", element.AriaRequired?.ToString() ?? ""),
                new("ARIA Has Popup", element.AriaHasPopup ?? ""),
                new("ARIA Level", element.AriaLevel?.ToString() ?? ""),
                new("ARIA Value Text", element.AriaValueText ?? "")
            };
            // Add other aria-* attributes
            if (element.AriaAttributes?.Count > 0)
            {
                foreach (var attr in element.AriaAttributes.Take(10))
                    webARIAProps.Add(new(attr.Key, attr.Value));
            }
            dgWebARIA.ItemsSource = webARIAProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // 3.3 Layout / Box Model
            var webLayoutProps = new List<KeyValuePair<string, string>>
            {
                new("Client Width", element.ClientWidth > 0 ? element.ClientWidth.ToString() : ""),
                new("Client Height", element.ClientHeight > 0 ? element.ClientHeight.ToString() : ""),
                new("Offset Width", element.OffsetWidth > 0 ? element.OffsetWidth.ToString() : ""),
                new("Offset Height", element.OffsetHeight > 0 ? element.OffsetHeight.ToString() : ""),
                new("Scroll Width", element.ScrollWidth > 0 ? element.ScrollWidth.ToString() : ""),
                new("Scroll Height", element.ScrollHeight > 0 ? element.ScrollHeight.ToString() : ""),
                new("Scroll Top", element.ScrollTop > 0 ? element.ScrollTop.ToString() : ""),
                new("Scroll Left", element.ScrollLeft > 0 ? element.ScrollLeft.ToString() : ""),
                new("Box Model Margin", element.BoxModel_Margin ?? ""),
                new("Box Model Padding", element.BoxModel_Padding ?? ""),
                new("Box Model Border", element.BoxModel_Border ?? "")
            };
            dgWebLayout.ItemsSource = webLayoutProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // 3.4 CSS Computed Styles
            var webCSSProps = new List<KeyValuePair<string, string>>
            {
                new("display", element.Style_Display ?? ""),
                new("position", element.Style_Position ?? ""),
                new("visibility", element.Style_Visibility ?? ""),
                new("opacity", element.Style_Opacity ?? ""),
                new("color", element.Style_Color ?? ""),
                new("background-color", element.Style_BackgroundColor ?? ""),
                new("font-size", element.Style_FontSize ?? ""),
                new("font-weight", element.Style_FontWeight ?? ""),
                new("z-index", element.Style_ZIndex ?? ""),
                new("pointer-events", element.Style_PointerEvents ?? ""),
                new("overflow", element.Style_Overflow ?? ""),
                new("transform", element.Style_Transform ?? "")
            };
            // Add other computed styles
            if (element.ComputedStyles?.Count > 0)
            {
                foreach (var style in element.ComputedStyles.Take(20))
                {
                    if (!webCSSProps.Any(p => p.Key == style.Key))
                        webCSSProps.Add(new(style.Key, style.Value));
                }
            }
            dgWebCSS.ItemsSource = webCSSProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // === 4. MSHTML / IHTMLDocument PROPERTIES ===
            var mshtmlElementProps = new List<KeyValuePair<string, string>>
            {
                // 4.1 IHTMLElement
                new("Parent Element Tag", element.MSHTML_ParentElementTag ?? ""),
                new("Children Count", element.MSHTML_ChildrenCount?.ToString() ?? ""),
                new("Outer Text", TruncateString(element.MSHTML_OuterText, 100)),
                new("Language", element.MSHTML_Language ?? ""),
                new("Lang", element.MSHTML_Lang ?? ""),
                new("Source Index", element.MSHTML_SourceIndex?.ToString() ?? ""),
                new("Scope Name", element.MSHTML_ScopeName ?? ""),
                new("Can Have Children", element.MSHTML_CanHaveChildren?.ToString() ?? ""),
                new("Can Have HTML", element.MSHTML_CanHaveHTML?.ToString() ?? ""),
                new("Is Content Editable", element.MSHTML_IsContentEditable?.ToString() ?? ""),
                new("Content Editable", element.MSHTML_ContentEditable ?? ""),
                new("Tab Index", element.MSHTML_TabIndex?.ToString() ?? ""),
                new("Access Key", element.MSHTML_AccessKey ?? ""),
                // 4.2 IHTMLElement2
                new("Client Left", element.MSHTML_ClientLeft?.ToString() ?? ""),
                new("Client Top", element.MSHTML_ClientTop?.ToString() ?? ""),
                new("Current Style", element.MSHTML_CurrentStyle ?? ""),
                new("Runtime Style", element.MSHTML_RuntimeStyle ?? ""),
                new("Ready State", element.MSHTML_ReadyState ?? ""),
                new("Dir (Text Direction)", element.MSHTML_Dir ?? ""),
                new("Scroll Left", element.MSHTML_ScrollLeft?.ToString() ?? ""),
                new("Scroll Top", element.MSHTML_ScrollTop?.ToString() ?? ""),
                new("Scroll Width", element.MSHTML_ScrollWidth?.ToString() ?? ""),
                new("Scroll Height", element.MSHTML_ScrollHeight?.ToString() ?? ""),
                // 4.3 IHTMLInputElement
                new("Default Value", element.MSHTML_DefaultValue ?? ""),
                new("Default Checked", element.MSHTML_DefaultChecked?.ToString() ?? ""),
                new("Form ID", element.MSHTML_Form_Id ?? ""),
                new("Form Name", element.MSHTML_Form_Name ?? ""),
                new("Form Action", element.MSHTML_Form_Action ?? ""),
                new("Form Method", element.MSHTML_Form_Method ?? ""),
                new("Max Length", element.MSHTML_MaxLength?.ToString() ?? ""),
                new("Size", element.MSHTML_Size?.ToString() ?? ""),
                new("Read Only", element.MSHTML_ReadOnly?.ToString() ?? ""),
                new("Accept", element.MSHTML_Accept ?? ""),
                new("Align", element.MSHTML_Align ?? ""),
                new("Indeterminate", element.MSHTML_IndeterminateState?.ToString() ?? ""),
                // 4.4 IHTMLSelectElement
                new("Selected Index", element.MSHTML_SelectedIndex?.ToString() ?? ""),
                new("Selected Value", element.MSHTML_SelectedValue ?? ""),
                new("Selected Text", element.MSHTML_SelectedText ?? ""),
                new("Options Length", element.MSHTML_OptionsLength?.ToString() ?? ""),
                new("Multiple", element.MSHTML_Multiple?.ToString() ?? ""),
                new("Options", element.MSHTML_Options?.Count > 0 ? string.Join(", ", element.MSHTML_Options.Take(10)) : ""),
                new("Option Values", element.MSHTML_OptionValues?.Count > 0 ? string.Join(", ", element.MSHTML_OptionValues.Take(10)) : ""),
                // 4.5 IHTMLTextAreaElement
                new("Cols", element.MSHTML_Cols?.ToString() ?? ""),
                new("Rows", element.MSHTML_Rows?.ToString() ?? ""),
                new("Wrap", element.MSHTML_Wrap ?? ""),
                // 4.6 IHTMLButtonElement
                new("Button Type", element.MSHTML_ButtonType ?? ""),
                new("Button Form Action", element.MSHTML_FormAction ?? ""),
                new("Button Form Method", element.MSHTML_FormMethod ?? ""),
                // 4.7 IHTMLAnchorElement
                new("Target", element.MSHTML_Target ?? ""),
                new("Protocol", element.MSHTML_Protocol ?? ""),
                new("Host", element.MSHTML_Host ?? ""),
                new("Hostname", element.MSHTML_Hostname ?? ""),
                new("Port", element.MSHTML_Port ?? ""),
                new("Pathname", element.MSHTML_Pathname ?? ""),
                new("Search (Query)", element.MSHTML_Search ?? ""),
                new("Hash (Fragment)", element.MSHTML_Hash ?? ""),
                new("Rel", element.MSHTML_Rel ?? ""),
                // 4.8 IHTMLImageElement
                new("Is Map", element.MSHTML_IsMap?.ToString() ?? ""),
                new("Natural Width", element.MSHTML_NaturalWidth?.ToString() ?? ""),
                new("Natural Height", element.MSHTML_NaturalHeight?.ToString() ?? ""),
                new("Complete", element.MSHTML_Complete?.ToString() ?? ""),
                new("Long Desc", element.MSHTML_LongDesc ?? ""),
                // 4.9-4.11 Table/Row/Cell
                new("Table Caption", element.MSHTML_Table_Caption ?? ""),
                new("Table Summary", element.MSHTML_Table_Summary ?? ""),
                new("Table Border", element.MSHTML_Table_Border ?? ""),
                new("Cell Padding", element.MSHTML_Table_CellPadding ?? ""),
                new("Cell Spacing", element.MSHTML_Table_CellSpacing ?? ""),
                new("Table Width", element.MSHTML_Table_Width ?? ""),
                new("Table Rows Count", element.MSHTML_Table_RowsCount?.ToString() ?? ""),
                new("Row Index", element.MSHTML_Row_RowIndex?.ToString() ?? ""),
                new("Section Row Index", element.MSHTML_Row_SectionRowIndex?.ToString() ?? ""),
                new("Cells Count", element.MSHTML_Row_CellsCount?.ToString() ?? ""),
                new("Cell Index", element.MSHTML_Cell_CellIndex?.ToString() ?? ""),
                new("Cell Scope", element.MSHTML_Cell_Scope ?? ""),
                new("Cell Headers", element.MSHTML_Cell_Headers ?? ""),
                // 4.12 IHTMLFrameElement
                new("Frame Src", element.MSHTML_Frame_Src ?? ""),
                new("Frame Name", element.MSHTML_Frame_Name ?? ""),
                new("Frame Scrolling", element.MSHTML_Frame_Scrolling ?? ""),
                new("Frame Border", element.MSHTML_Frame_FrameBorder ?? ""),
                new("No Resize", element.MSHTML_Frame_NoResize?.ToString() ?? "")
            };
            dgMSHTMLElement.ItemsSource = mshtmlElementProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            var mshtmlDocProps = new List<KeyValuePair<string, string>>
            {
                // 4.13 IHTMLDocument
                new("Document Title", element.DocumentTitle ?? ""),
                new("Document URL", element.DocumentUrl ?? ""),
                new("Document Domain", element.DocumentDomain ?? ""),
                new("Document Ready State", element.DocumentReadyState ?? ""),
                new("Document Charset", element.DocumentCharset ?? ""),
                new("Document Last Modified", element.DocumentLastModified ?? ""),
                new("Document Referrer", element.DocumentReferrer ?? ""),
                new("Document Compat Mode", element.DocumentCompatMode ?? ""),
                new("Document Design Mode", element.DocumentDesignMode ?? ""),
                new("Document DocType", element.DocumentDocType ?? ""),
                new("Document Dir", element.DocumentDir ?? ""),
                new("Document Cookie", TruncateString(element.DocumentCookie, 100)),
                new("Frames Count", element.DocumentFramesCount?.ToString() ?? ""),
                new("Scripts Count", element.DocumentScriptsCount?.ToString() ?? ""),
                new("Links Count", element.DocumentLinksCount?.ToString() ?? ""),
                new("Images Count", element.DocumentImagesCount?.ToString() ?? ""),
                new("Forms Count", element.DocumentFormsCount?.ToString() ?? ""),
                new("Active Element", element.DocumentActiveElement ?? ""),
                // 4.14 IHTMLDocument2/3/4/5
                new("Protocol", element.MSHTML_Doc_Protocol ?? ""),
                new("Name Prop", element.MSHTML_Doc_NameProp ?? ""),
                new("File Created Date", element.MSHTML_Doc_FileCreatedDate ?? ""),
                new("File Modified Date", element.MSHTML_Doc_FileModifiedDate ?? ""),
                new("File Size", element.MSHTML_Doc_FileSize ?? ""),
                new("Mime Type", element.MSHTML_Doc_MimeType ?? ""),
                new("Security", element.MSHTML_Doc_Security ?? ""),
                new("Anchors Count", element.MSHTML_Doc_Anchors_Count?.ToString() ?? ""),
                new("Applets Count", element.MSHTML_Doc_Applets_Count?.ToString() ?? ""),
                new("Embeds Count", element.MSHTML_Doc_Embeds_Count?.ToString() ?? ""),
                new("All Elements Count", element.MSHTML_Doc_All_Count?.ToString() ?? "")
            };
            dgMSHTMLDocument.ItemsSource = mshtmlDocProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // === 5. WIN32 API / SPY++ PROPERTIES ===
            var win32HandleProps = new List<KeyValuePair<string, string>>
            {
                new("HWND", element.Win32_HWND != IntPtr.Zero ? $"0x{element.Win32_HWND.ToInt64():X8}" : ""),
                new("Parent HWND", element.Win32_ParentHWND != IntPtr.Zero ? $"0x{element.Win32_ParentHWND.ToInt64():X8}" : ""),
                new("Win32 Class Name", element.Win32_ClassName ?? ""),
                new("Win32 Window Text", element.Win32_WindowText ?? ""),
                new("Thread Id", element.Win32_ThreadId > 0 ? element.Win32_ThreadId.ToString() : ""),
                new("Control Id", element.Win32_ControlId ?? ""),
                new("Is Unicode", element.Win32_IsUnicode?.ToString() ?? ""),
                new("Is Visible", element.Win32_IsVisible?.ToString() ?? ""),
                new("Is Enabled", element.Win32_IsEnabled?.ToString() ?? ""),
                new("Is Maximized", element.Win32_IsMaximized?.ToString() ?? ""),
                new("Is Minimized", element.Win32_IsMinimized?.ToString() ?? ""),
                new("Instance Handle", element.Win32_InstanceHandle != IntPtr.Zero ? $"0x{element.Win32_InstanceHandle.ToInt64():X8}" : ""),
                new("Menu Handle", element.Win32_MenuHandle != IntPtr.Zero ? $"0x{element.Win32_MenuHandle.ToInt64():X8}" : ""),
                new("WndProc", element.Win32_WndProc != IntPtr.Zero ? $"0x{element.Win32_WndProc.ToInt64():X8}" : "")
            };
            dgWin32Handle.ItemsSource = win32HandleProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            var win32StyleProps = new List<KeyValuePair<string, string>>
            {
                new("Window Styles (Raw)", element.Win32_WindowStyles > 0 ? $"0x{element.Win32_WindowStyles:X8}" : ""),
                new("Window Styles (Parsed)", element.Win32_WindowStyles_Parsed ?? ""),
                new("Extended Styles (Raw)", element.Win32_ExtendedStyles > 0 ? $"0x{element.Win32_ExtendedStyles:X8}" : ""),
                new("Extended Styles (Parsed)", element.Win32_ExtendedStyles_Parsed ?? "")
            };
            dgWin32Styles.ItemsSource = win32StyleProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            var win32RectProps = new List<KeyValuePair<string, string>>
            {
                new("Window Rect", element.Win32_WindowRect.Width > 0 ? $"X:{element.Win32_WindowRect.X}, Y:{element.Win32_WindowRect.Y}, W:{element.Win32_WindowRect.Width}, H:{element.Win32_WindowRect.Height}" : ""),
                new("Client Rect", element.Win32_ClientRect.Width > 0 ? $"W:{element.Win32_ClientRect.Width}, H:{element.Win32_ClientRect.Height}" : "")
            };
            dgWin32Rect.ItemsSource = win32RectProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();

            // === 6. POSITION & SIZE ===
            var positionProps = new List<KeyValuePair<string, string>>
            {
                new("Bounding Rectangle", element.BoundingRectangle.Width > 0 ? element.BoundingRectangle.ToString() : ""),
                new("X", element.X.ToString()),
                new("Y", element.Y.ToString()),
                new("Width", element.Width.ToString()),
                new("Height", element.Height.ToString()),
                new("Clickable Point", element.ClickablePoint.X > 0 || element.ClickablePoint.Y > 0 ? $"{element.ClickablePoint.X}, {element.ClickablePoint.Y}" : ""),
                new("Client Rect", element.ClientRect.Width > 0 ? element.ClientRect.ToString() : "")
            };
            dgPosition.ItemsSource = positionProps.Where(p => !string.IsNullOrEmpty(p.Value) && p.Value != "0").ToList();

            // === 7. SELECTORS - Update TextBoxes ===
            txtFullXPath.Text = element.FullXPath ?? "N/A";
            txtPlaywrightSelector.Text = element.PlaywrightSelector ?? "N/A";
            txtTreePath.Text = element.TreePath ?? element.ElementPath ?? "N/A";

            // === 8. HIERARCHY ===
            var hierarchyProps = new List<KeyValuePair<string, string>>
            {
                new("Parent Name", element.ParentName ?? ""),
                new("Parent Id", element.ParentId ?? ""),
                new("Parent Class Name", element.ParentClassName ?? ""),
                new("Children Count", element.Children?.Count.ToString() ?? "0"),
                new("Child Index", element.ChildIndex.ToString()),
                new("Tree Level", element.TreeLevel.ToString()),
                new("Owner Window", element.OwnerWindow ?? ""),
                new("Control Container", element.ControlContainer ?? ""),
                new("Window Title", element.WindowTitle ?? ""),
                new("Window Class Name", element.WindowClassName ?? ""),
                new("Application Name", element.ApplicationName ?? ""),
                new("Application Path", element.ApplicationPath ?? "")
            };
            dgHierarchy.ItemsSource = hierarchyProps.Where(p => !string.IsNullOrEmpty(p.Value) && p.Value != "0").ToList();

            // === 9. TABLE/GRID PROPERTIES ===
            var tableProps = new List<KeyValuePair<string, string>>
            {
                new("Is Table Cell", element.IsTableCell.ToString()),
                new("Is Table Header", element.IsTableHeader.ToString()),
                new("Row Index", element.RowIndex >= 0 ? element.RowIndex.ToString() : ""),
                new("Column Index", element.ColumnIndex >= 0 ? element.ColumnIndex.ToString() : ""),
                new("Row Count", element.RowCount >= 0 ? element.RowCount.ToString() : ""),
                new("Column Count", element.ColumnCount >= 0 ? element.ColumnCount.ToString() : ""),
                new("Row Span", element.RowSpan > 1 ? element.RowSpan.ToString() : ""),
                new("Column Span", element.ColumnSpan > 1 ? element.ColumnSpan.ToString() : ""),
                new("Table Name", element.TableName ?? ""),
                new("Column Headers", element.ColumnHeaders?.Count > 0 ? string.Join(", ", element.ColumnHeaders) : ""),
                new("Row Headers", element.RowHeaders?.Count > 0 ? string.Join(", ", element.RowHeaders) : ""),
                new("Row Or Column Major", element.Table_RowOrColumnMajor ?? "")
            };
            dgTable.ItemsSource = tableProps.Where(p => !string.IsNullOrEmpty(p.Value) && p.Value != "False" && p.Value != "-1").ToList();

            // === 10. STATE PROPERTIES ===
            var stateProps = new List<KeyValuePair<string, string>>
            {
                new("Is Visible", element.IsVisible.ToString()),
                new("Is Hidden", element.IsHidden.ToString()),
                new("Is Checked", element.IsChecked.ToString()),
                new("Is Disabled", element.IsDisabled.ToString()),
                new("Is Editable", element.IsEditable.ToString()),
                new("Is Selected", element.IsSelected.ToString()),
                new("Is Focused", element.IsFocused.ToString()),
                new("Is Expanded", element.IsExpanded.ToString())
            };
            dgState.ItemsSource = stateProps.ToList();

            // === 11. CUSTOM PROPERTIES ===
            var customProps = new List<KeyValuePair<string, string>>();
            if (element.CustomProperties?.Count > 0)
            {
                foreach (var prop in element.CustomProperties)
                {
                    customProps.Add(new(prop.Key, prop.Value?.ToString() ?? ""));
                }
            }
            // Add collection info
            customProps.Add(new("Collection Duration", $"{element.CollectionDuration.TotalMilliseconds}ms"));
            if (element.CollectionErrors?.Count > 0)
            {
                customProps.Add(new("Collection Errors", string.Join("; ", element.CollectionErrors.Take(5))));
            }
            dgCustom.ItemsSource = customProps.Where(p => !string.IsNullOrEmpty(p.Value)).ToList();
        }

        private string TruncateString(string value, int maxLength)
        {
            if (string.IsNullOrEmpty(value)) return "";
            if (value.Length <= maxLength) return value;
            return value.Substring(0, maxLength) + "...";
        }

        private void ClearElementDetails()
        {
            txtRawProperties.Clear();

            // Clear all DataGrids
            dgBasicProperties.ItemsSource = null;
            dgUIABasic.ItemsSource = null;
            dgUIAPatterns.ItemsSource = null;
            dgUIALegacy.ItemsSource = null;
            dgWebDOM.ItemsSource = null;
            dgWebARIA.ItemsSource = null;
            dgWebLayout.ItemsSource = null;
            dgWebCSS.ItemsSource = null;
            dgMSHTMLElement.ItemsSource = null;
            dgMSHTMLDocument.ItemsSource = null;
            dgWin32Handle.ItemsSource = null;
            dgWin32Styles.ItemsSource = null;
            dgWin32Rect.ItemsSource = null;
            dgPosition.ItemsSource = null;
            dgHierarchy.ItemsSource = null;
            dgTable.ItemsSource = null;
            dgState.ItemsSource = null;
            dgCustom.ItemsSource = null;

            // Clear TextBoxes
            txtXPath.Clear();
            txtFullXPath.Clear();
            txtCssSelector.Clear();
            txtWindowsPath.Clear();
            txtPlaywrightSelector.Clear();
            txtTreePath.Clear();
            txtSourceCode.Clear();
            imgScreenshot.Source = null;
        }

        #endregion

        #region Helper Methods

        private void HideMainWindowShowFloating()
        {
            // Hide main window
            this.WindowState = WindowState.Minimized;
            this.Hide();

            // Update and show floating window
            var modeText = ((ComboBoxItem)cmbInspectionMode.SelectedItem).Content.ToString();
            _floatingWindow.UpdateMode(modeText);
            _floatingWindow.UpdateStatus("Waiting for input...");
            _floatingWindow.UpdateElementCount(_collectedElements.Count);
            _floatingWindow.Show();
        }

        private void ShowMainWindow()
        {
            // Hide floating window
            _floatingWindow.Hide();

            // Show main window
            this.Show();
            this.WindowState = WindowState.Normal;
            this.Activate();
            this.Focus();
        }

        private void UpdateFloatingWindow(string status, int elementCount)
        {
            if (_floatingWindow.IsVisible)
            {
                _floatingWindow.UpdateStatus(status);
                _floatingWindow.UpdateElementCount(elementCount);
            }
        }

        private void Topmost_Click(object sender, RoutedEventArgs e)
        {
            this.Topmost = btnTopmost.IsChecked == true;

            // Update button appearance based on state
            if (this.Topmost)
            {
                btnTopmost.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(76, 175, 80)); // Green when active
                LogToConsole("Her Zaman Üstte: AÇIK");
            }
            else
            {
                btnTopmost.Background = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 87, 34)); // Orange when inactive
                LogToConsole("Her Zaman Üstte: KAPALI");
            }
        }

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
                // NOTE: Playwright is excluded from Auto Detect to prevent automatic browser launching
                var point = System.Windows.Forms.Cursor.Position;
                var wpfPoint = new System.Windows.Point(point.X, point.Y);
                return _detectors.FirstOrDefault(d => d.Name != "Playwright" && d.CanDetect(wpfPoint));
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
            else if (techIndex == 4) // Playwright
            {
                // Playwright selected - initialize browser if needed
                var playwrightDetector = _detectors.FirstOrDefault(d => d.Name == "Playwright") as PlaywrightDetector;
                if (playwrightDetector != null)
                {
                    // Initialize browser synchronously when user explicitly selects Playwright
                    Task.Run(async () => await playwrightDetector.EnsureInitializedAsync()).GetAwaiter().GetResult();
                }
                return playwrightDetector;
            }
            else if (techIndex == 5) // All Technologies
            {
                // Return the first detector that can detect
                // NOTE: Playwright is excluded from All Technologies to prevent automatic browser launching
                var point = System.Windows.Forms.Cursor.Position;
                var wpfPoint = new System.Windows.Point(point.X, point.Y);
                return _detectors.FirstOrDefault(d => d.Name != "Playwright" && d.CanDetect(wpfPoint)) ?? _detectors.FirstOrDefault(d => d.Name != "Playwright");
            }

            return _detectors.FirstOrDefault();
        }

        private void LogToConsole(string message, Core.Utils.LogLevel level = Core.Utils.LogLevel.Info)
        {
            // Log to file
            _logger?.Log(message, level);

            // Log to UI console
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
            try
            {
                // Skip if UI not fully loaded
                if (txtTreeSearch == null || txtConsole == null) return;

                var searchText = txtTreeSearch.Text;

                // Ignore placeholder text
                if (string.IsNullOrWhiteSpace(searchText) || searchText == "Search elements..." || searchText == "Search...")
                {
                    _searchResults = null;
                    _currentSearchIndex = -1;
                    return;
                }

                // Perform search
                PerformSearch(searchText);
            }
            catch (Exception ex)
            {
                // Silently ignore during startup
                System.Diagnostics.Debug.WriteLine($"Search error: {ex.Message}");
            }
        }

        private void TreeSearch_GotFocus(object sender, RoutedEventArgs e)
        {
            if (txtTreeSearch.Text == "Search elements..." || txtTreeSearch.Text == "Search...")
            {
                txtTreeSearch.Text = "";
                txtTreeSearch.Foreground = System.Windows.Media.Brushes.Black;
            }
        }

        private void TreeSearch_LostFocus(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtTreeSearch.Text))
            {
                txtTreeSearch.Text = "Search...";
                txtTreeSearch.Foreground = System.Windows.Media.Brushes.Gray;
            }
        }

        private void PerformSearch(string searchText)
        {
            try
            {
                searchText = searchText.ToLower();
                _searchResults = new List<ElementInfo>();
                _currentSearchIndex = -1;

                // Search through all collected elements
                foreach (var element in _collectedElements)
                {
                    if (ElementMatchesSearch(element, searchText))
                    {
                        _searchResults.Add(element);
                    }

                    // Search children recursively
                    SearchElementChildren(element, searchText);
                }

                // Show results and navigate to first match
                if (_searchResults.Count > 0)
                {
                    _currentSearchIndex = 0;
                    SelectElementInTree(_searchResults[0]);
                    LogToConsole($"Found {_searchResults.Count} matching elements", Core.Utils.LogLevel.Info);
                }
                else
                {
                    LogToConsole("No matching elements found", Core.Utils.LogLevel.Warning);
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Search error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private void SearchElementChildren(ElementInfo parent, string searchText)
        {
            if (parent.Children == null || parent.Children.Count == 0)
                return;

            foreach (var child in parent.Children)
            {
                if (ElementMatchesSearch(child, searchText))
                {
                    _searchResults.Add(child);
                }

                // Recursively search children
                SearchElementChildren(child, searchText);
            }
        }

        private bool ElementMatchesSearch(ElementInfo element, string searchText)
        {
            if (element == null)
                return false;

            // Search in common properties
            return (element.Name?.ToLower().Contains(searchText) ?? false) ||
                   (element.AutomationId?.ToLower().Contains(searchText) ?? false) ||
                   (element.ClassName?.ToLower().Contains(searchText) ?? false) ||
                   (element.ControlType?.ToLower().Contains(searchText) ?? false) ||
                   (element.ElementType?.ToLower().Contains(searchText) ?? false) ||
                   (element.TagName?.ToLower().Contains(searchText) ?? false) ||
                   (element.HtmlId?.ToLower().Contains(searchText) ?? false) ||
                   (element.InnerText?.ToLower().Contains(searchText) ?? false) ||
                   (element.Value?.ToLower().Contains(searchText) ?? false) ||
                   (element.Description?.ToLower().Contains(searchText) ?? false);
        }

        private void SelectElementInTree(ElementInfo element)
        {
            try
            {
                // Find the element in the TreeView and select it
                foreach (var item in tvElements.Items)
                {
                    if (item is ElementInfo rootElement)
                    {
                        if (rootElement.Id == element.Id)
                        {
                            // Select root element
                            var container = tvElements.ItemContainerGenerator.ContainerFromItem(rootElement) as TreeViewItem;
                            if (container != null)
                            {
                                container.IsSelected = true;
                                container.BringIntoView();
                            }
                            return;
                        }

                        // Search in children
                        if (SelectChildInTree(rootElement, element))
                            return;
                    }
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Tree selection error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private bool SelectChildInTree(ElementInfo parent, ElementInfo target)
        {
            if (parent.Children == null || parent.Children.Count == 0)
                return false;

            foreach (var child in parent.Children)
            {
                if (child.Id == target.Id)
                {
                    // Found the target - need to expand parent and select child
                    // This is simplified - full implementation would need TreeViewItem navigation
                    return true;
                }

                if (SelectChildInTree(child, target))
                    return true;
            }

            return false;
        }

        private void TreeSearch_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Navigate to next search result
                if (_searchResults == null || _searchResults.Count == 0)
                {
                    LogToConsole("No search results available", Core.Utils.LogLevel.Warning);
                    return;
                }

                // Move to next result (wrap around)
                _currentSearchIndex++;
                if (_currentSearchIndex >= _searchResults.Count)
                {
                    _currentSearchIndex = 0;
                }

                // Select the current result
                SelectElementInTree(_searchResults[_currentSearchIndex]);
                LogToConsole($"Showing result {_currentSearchIndex + 1} of {_searchResults.Count}", Core.Utils.LogLevel.Info);
            }
            catch (Exception ex)
            {
                LogToConsole($"Search navigation error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private void ExpandAll_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement expand all functionality
        }

        private void CollapseAll_Click(object sender, RoutedEventArgs e)
        {
            // TODO: Implement collapse all functionality
        }

        private async void Refresh_Click(object sender, RoutedEventArgs e)
        {
            if (_currentElement == null)
            {
                LogToConsole("No element selected to refresh", Core.Utils.LogLevel.Warning);
                return;
            }

            try
            {
                LogToConsole("Refreshing current element...", Core.Utils.LogLevel.Info);

                // Get current collection profile
                var profile = GetSelectedProfile();

                // Refresh element from each detector that originally detected it
                var refreshedElements = new List<ElementInfo>();

                foreach (var detector in _detectors)
                {
                    try
                    {
                        var refreshed = await detector.RefreshElement(_currentElement, profile);
                        if (refreshed != null && refreshed.Id != _currentElement.Id)
                        {
                            // Successfully refreshed with new data
                            refreshedElements.Add(refreshed);
                            LogToConsole($"Refreshed element using {detector.Name}", Core.Utils.LogLevel.Info);
                        }
                    }
                    catch (Exception ex)
                    {
                        LogToConsole($"{detector.Name} refresh error: {ex.Message}", Core.Utils.LogLevel.Error);
                    }
                }

                if (refreshedElements.Count > 0)
                {
                    // Merge all refreshed data
                    var mergedElement = MergeElementInfo(refreshedElements);

                    // Update the current element with refreshed data
                    _currentElement = mergedElement;
                    _currentElement.CaptureTime = DateTime.Now;

                    // Update UI
                    DisplayElementInfo(_currentElement);
                    LogToConsole($"Element refreshed successfully with {refreshedElements.Count} detector(s)", Core.Utils.LogLevel.Info);
                }
                else
                {
                    LogToConsole("Could not refresh element - no detectors returned updated data", Core.Utils.LogLevel.Warning);
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Refresh error: {ex.Message}", Core.Utils.LogLevel.Error);
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

        private async void ImportSession_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Create OpenFileDialog
                var openFileDialog = new Microsoft.Win32.OpenFileDialog
                {
                    Title = "Import Session",
                    Filter = "JSON Files (*.json)|*.json|XML Files (*.xml)|*.xml|All Files (*.*)|*.*",
                    DefaultExt = ".json",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    var filePath = openFileDialog.FileName;
                    var extension = System.IO.Path.GetExtension(filePath).ToLower();

                    LogToConsole($"Importing session from: {filePath}", Core.Utils.LogLevel.Info);

                    List<ElementInfo> importedElements = null;

                    // Import based on file extension
                    switch (extension)
                    {
                        case ".json":
                            importedElements = await ImportFromJson(filePath);
                            break;
                        case ".xml":
                            importedElements = await ImportFromXml(filePath);
                            break;
                        default:
                            LogToConsole($"Unsupported file format: {extension}", Core.Utils.LogLevel.Warning);
                            System.Windows.MessageBox.Show($"Unsupported file format: {extension}\nPlease select a JSON or XML file.",
                                "Import Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                    }

                    if (importedElements != null && importedElements.Count > 0)
                    {
                        // Clear existing elements
                        _collectedElements.Clear();

                        // Add imported elements
                        foreach (var element in importedElements)
                        {
                            _collectedElements.Add(element);
                        }

                        // Bind to TreeView
                        tvElements.ItemsSource = null;
                        tvElements.ItemsSource = _collectedElements;

                        // Update element count
                        txtElementCount.Text = _collectedElements.Count.ToString();

                        LogToConsole($"Successfully imported {importedElements.Count} elements", Core.Utils.LogLevel.Info);
                        _logger?.LogInfo($"Imported {importedElements.Count} elements from: {filePath}");

                        System.Windows.MessageBox.Show($"Successfully imported {importedElements.Count} elements from:\n{filePath}",
                            "Import Successful", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    else
                    {
                        LogToConsole("No elements found in the imported file", Core.Utils.LogLevel.Warning);
                        System.Windows.MessageBox.Show("No elements found in the imported file.",
                            "Import Warning", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Import error: {ex.Message}", Core.Utils.LogLevel.Error);
                _logger?.LogError($"Import error: {ex.Message}\n{ex.StackTrace}");
                System.Windows.MessageBox.Show($"Failed to import session:\n{ex.Message}",
                    "Import Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async Task<List<ElementInfo>> ImportFromJson(string filePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var jsonContent = File.ReadAllText(filePath);
                    var options = new System.Text.Json.JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true
                    };

                    // Try to parse as export format first (with ExportDate, TotalElements, Elements)
                    try
                    {
                        var exportData = System.Text.Json.JsonSerializer.Deserialize<JsonExportFormat>(jsonContent, options);
                        if (exportData?.Elements != null)
                        {
                            return exportData.Elements;
                        }
                    }
                    catch
                    {
                        // If that fails, try to parse as a direct list of ElementInfo
                        var elements = System.Text.Json.JsonSerializer.Deserialize<List<ElementInfo>>(jsonContent, options);
                        if (elements != null)
                        {
                            return elements;
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"JSON parsing error: {ex.Message}", ex);
                }

                return new List<ElementInfo>();
            });
        }

        private async Task<List<ElementInfo>> ImportFromXml(string filePath)
        {
            return await Task.Run(() =>
            {
                try
                {
                    var xmlContent = File.ReadAllText(filePath);
                    var serializer = new System.Xml.Serialization.XmlSerializer(typeof(List<ElementInfo>));

                    using (var reader = new System.IO.StringReader(xmlContent))
                    {
                        var elements = serializer.Deserialize(reader) as List<ElementInfo>;
                        return elements ?? new List<ElementInfo>();
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"XML parsing error: {ex.Message}", ex);
                }
            });
        }

        // Helper class for JSON deserialization
        private class JsonExportFormat
        {
            public DateTime ExportDate { get; set; }
            public int TotalElements { get; set; }
            public List<ElementInfo> Elements { get; set; }
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
F1 - Start Inspection (Hide Main Window)
F2 - Stop Inspection (Show Main Window)
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

        /// <summary>
        /// Copies a comprehensive report with ALL properties from all 5 technologies to clipboard
        /// </summary>
        private void CopyFullReport_Click(object sender, RoutedEventArgs e)
        {
            if (_currentElement == null)
            {
                LogToConsole("No element selected. Please inspect an element first.");
                System.Windows.MessageBox.Show(
                    "No element selected.\n\nPlease use 'Start' to inspect an element first.",
                    "No Element Selected",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
                return;
            }

            try
            {
                // Generate full comprehensive report with all 5 technologies
                var fullReport = GenerateFullReport(_currentElement);

                // Copy to clipboard
                System.Windows.Clipboard.SetText(fullReport);

                // Calculate approximate property count
                int propertyCount = CountProperties(_currentElement);

                LogToConsole($"✅ Full report copied to clipboard!");
                LogToConsole($"   Technologies: {string.Join(", ", _currentElement.TechnologiesUsed ?? new List<string> { "Unknown" })}");
                LogToConsole($"   Properties: ~{propertyCount} items");
                LogToConsole($"   Report size: {fullReport.Length:N0} characters");

                // Show success message
                System.Windows.MessageBox.Show(
                    $"Full report copied to clipboard!\n\n" +
                    $"Technologies: {string.Join(", ", _currentElement.TechnologiesUsed ?? new List<string> { "N/A" })}\n" +
                    $"Properties: ~{propertyCount} items\n" +
                    $"Report size: {fullReport.Length:N0} characters\n\n" +
                    $"You can now paste (Ctrl+V) anywhere.",
                    "Copy Successful",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                LogToConsole($"❌ Error copying report: {ex.Message}");
                System.Windows.MessageBox.Show(
                    $"Error copying report to clipboard:\n{ex.Message}",
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Counts the number of non-null properties in an ElementInfo
        /// </summary>
        private int CountProperties(ElementInfo element)
        {
            if (element == null) return 0;

            int count = 0;
            var properties = typeof(ElementInfo).GetProperties();

            foreach (var prop in properties)
            {
                try
                {
                    var value = prop.GetValue(element);
                    if (value != null)
                    {
                        // Handle different types
                        if (value is string str && !string.IsNullOrEmpty(str))
                            count++;
                        else if (value is System.Collections.ICollection collection && collection.Count > 0)
                            count += collection.Count;
                        else if (value is bool || value is int || value is double || value is IntPtr)
                        {
                            // Check if it's not a default value
                            if (value is IntPtr ptr && ptr != IntPtr.Zero)
                                count++;
                            else if (value is int i && i != 0)
                                count++;
                            else if (value is double d && d != 0)
                                count++;
                            else if (value is bool)
                                count++;
                        }
                        else if (value is Nullable<bool> || value is Nullable<int> || value is Nullable<double>)
                        {
                            count++;
                        }
                    }
                }
                catch { }
            }

            return count;
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

        private async void ExportWithDialog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Create SaveFileDialog
                var saveFileDialog = new Microsoft.Win32.SaveFileDialog
                {
                    Title = "Export UI Elements Data",
                    Filter = "Text Files (*.txt)|*.txt|JSON Files (*.json)|*.json|CSV Files (*.csv)|*.csv|XML Files (*.xml)|*.xml|All Files (*.*)|*.*",
                    FilterIndex = 1,
                    FileName = $"UIElements_{DateTime.Now:yyyy-MM-dd_HH-mm-ss}",
                    InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop)
                };

                // Show dialog and get user's choice
                if (saveFileDialog.ShowDialog() == true)
                {
                    var filePath = saveFileDialog.FileName;
                    var extension = System.IO.Path.GetExtension(filePath).ToLower();

                    // Generate content based on file type
                    string content;
                    switch (extension)
                    {
                        case ".json":
                            content = await GenerateJsonExport();
                            break;
                        case ".csv":
                            content = await GenerateCsvExport();
                            break;
                        case ".xml":
                            content = await GenerateXmlExport();
                            break;
                        case ".txt":
                        default:
                            content = GenerateTextExport();
                            break;
                    }

                    await File.WriteAllTextAsync(filePath, content);

                    LogToConsole($"Exported to: {filePath}");
                    _logger?.LogInfo($"Exported {_collectedElements.Count} elements to: {filePath}");

                    System.Windows.MessageBox.Show($"Data exported successfully to:\n{filePath}", "Export Successful",
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    LogToConsole("Export cancelled by user.");
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Export error: {ex.Message}");
                _logger?.LogException(ex, "Export failed");
                System.Windows.MessageBox.Show($"Export failed: {ex.Message}", "Export Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private string GenerateTextExport()
        {
            var sb = new StringBuilder();
            sb.AppendLine($"UI Elements Export - {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"Total Elements: {_collectedElements.Count}");
            sb.AppendLine(new string('=', 80));
            sb.AppendLine();

            foreach (var element in _collectedElements)
            {
                sb.AppendLine(GenerateRawProperties(element));
                sb.AppendLine(new string('=', 80));
                sb.AppendLine();
            }

            return sb.ToString();
        }

        private async Task<string> GenerateJsonExport()
        {
            return await Task.Run(() =>
            {
                var export = new
                {
                    ExportDate = DateTime.Now,
                    TotalElements = _collectedElements.Count,
                    Elements = _collectedElements
                };
                return System.Text.Json.JsonSerializer.Serialize(export, new System.Text.Json.JsonSerializerOptions
                {
                    WriteIndented = true
                });
            });
        }

        private async Task<string> GenerateCsvExport()
        {
            return await Task.Run(() =>
            {
                var sb = new StringBuilder();
                // CSV Header - Added table properties
                sb.AppendLine("DetectionMethod,ElementType,Name,ClassName,AutomationId,ControlType,TagName,HtmlId,X,Y,Width,Height,TreePath,ElementPath," +
                    "RowIndex,ColumnIndex,RowCount,ColumnCount,RowSpan,ColumnSpan,IsTableCell,IsTableHeader,TableName,ColumnHeaders,RowHeaders," +
                    "XPath,CssSelector,PlaywrightTableSelector,Value,InnerText,IsVisible,IsEnabled");

                foreach (var element in _collectedElements)
                {
                    sb.AppendLine($"\"{EscapeCsv(element.DetectionMethod)}\"," +
                        $"\"{EscapeCsv(element.ElementType)}\"," +
                        $"\"{EscapeCsv(element.Name)}\"," +
                        $"\"{EscapeCsv(element.ClassName)}\"," +
                        $"\"{EscapeCsv(element.AutomationId)}\"," +
                        $"\"{EscapeCsv(element.ControlType)}\"," +
                        $"\"{EscapeCsv(element.TagName)}\"," +
                        $"\"{EscapeCsv(element.HtmlId)}\"," +
                        $"{element.X},{element.Y},{element.Width},{element.Height}," +
                        $"\"{EscapeCsv(element.TreePath)}\"," +
                        $"\"{EscapeCsv(element.ElementPath)}\"," +
                        $"{element.RowIndex},{element.ColumnIndex},{element.RowCount},{element.ColumnCount}," +
                        $"{element.RowSpan},{element.ColumnSpan}," +
                        $"{element.IsTableCell},{element.IsTableHeader}," +
                        $"\"{EscapeCsv(element.TableName)}\"," +
                        $"\"{EscapeCsv(string.Join(";", element.ColumnHeaders))}\"," +
                        $"\"{EscapeCsv(string.Join(";", element.RowHeaders))}\"," +
                        $"\"{EscapeCsv(element.XPath)}\"," +
                        $"\"{EscapeCsv(element.CssSelector)}\"," +
                        $"\"{EscapeCsv(element.PlaywrightTableSelector)}\"," +
                        $"\"{EscapeCsv(element.Value)}\"," +
                        $"\"{EscapeCsv(element.InnerText)}\"," +
                        $"{element.IsVisible},{element.IsEnabled}");
                }

                return sb.ToString();
            });
        }

        private string EscapeCsv(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";
            return value.Replace("\"", "\"\"").Replace("\n", " ").Replace("\r", "");
        }

        private async Task<string> GenerateXmlExport()
        {
            return await Task.Run(() =>
            {
                var sb = new StringBuilder();
                sb.AppendLine("<?xml version=\"1.0\" encoding=\"utf-8\"?>");
                sb.AppendLine("<UIElementsExport>");
                sb.AppendLine($"  <ExportDate>{DateTime.Now:yyyy-MM-dd HH:mm:ss}</ExportDate>");
                sb.AppendLine($"  <TotalElements>{_collectedElements.Count}</TotalElements>");
                sb.AppendLine("  <Elements>");

                foreach (var element in _collectedElements)
                {
                    sb.AppendLine("    <Element>");
                    sb.AppendLine($"      <DetectionMethod>{EscapeXml(element.DetectionMethod)}</DetectionMethod>");
                    sb.AppendLine($"      <ElementType>{EscapeXml(element.ElementType)}</ElementType>");
                    sb.AppendLine($"      <Name>{EscapeXml(element.Name)}</Name>");
                    sb.AppendLine($"      <ClassName>{EscapeXml(element.ClassName)}</ClassName>");
                    sb.AppendLine($"      <AutomationId>{EscapeXml(element.AutomationId)}</AutomationId>");
                    sb.AppendLine($"      <ControlType>{EscapeXml(element.ControlType)}</ControlType>");
                    sb.AppendLine($"      <TreePath>{EscapeXml(element.TreePath)}</TreePath>");
                    sb.AppendLine($"      <ElementPath>{EscapeXml(element.ElementPath)}</ElementPath>");
                    sb.AppendLine($"      <RowIndex>{element.RowIndex}</RowIndex>");
                    sb.AppendLine($"      <ColumnIndex>{element.ColumnIndex}</ColumnIndex>");
                    sb.AppendLine($"      <RowCount>{element.RowCount}</RowCount>");
                    sb.AppendLine($"      <ColumnCount>{element.ColumnCount}</ColumnCount>");
                    sb.AppendLine($"      <RowSpan>{element.RowSpan}</RowSpan>");
                    sb.AppendLine($"      <ColumnSpan>{element.ColumnSpan}</ColumnSpan>");
                    sb.AppendLine($"      <IsTableCell>{element.IsTableCell}</IsTableCell>");
                    sb.AppendLine($"      <IsTableHeader>{element.IsTableHeader}</IsTableHeader>");
                    sb.AppendLine($"      <TableName>{EscapeXml(element.TableName)}</TableName>");
                    if (element.ColumnHeaders.Count > 0)
                    {
                        sb.AppendLine("      <ColumnHeaders>");
                        foreach (var header in element.ColumnHeaders)
                            sb.AppendLine($"        <Header>{EscapeXml(header)}</Header>");
                        sb.AppendLine("      </ColumnHeaders>");
                    }
                    if (element.RowHeaders.Count > 0)
                    {
                        sb.AppendLine("      <RowHeaders>");
                        foreach (var header in element.RowHeaders)
                            sb.AppendLine($"        <Header>{EscapeXml(header)}</Header>");
                        sb.AppendLine("      </RowHeaders>");
                    }
                    sb.AppendLine($"      <XPath>{EscapeXml(element.XPath)}</XPath>");
                    sb.AppendLine($"      <CssSelector>{EscapeXml(element.CssSelector)}</CssSelector>");
                    sb.AppendLine($"      <PlaywrightTableSelector>{EscapeXml(element.PlaywrightTableSelector)}</PlaywrightTableSelector>");
                    sb.AppendLine($"      <Position X=\"{element.X}\" Y=\"{element.Y}\" Width=\"{element.Width}\" Height=\"{element.Height}\"/>");
                    sb.AppendLine($"      <IsVisible>{element.IsVisible}</IsVisible>");
                    sb.AppendLine($"      <IsEnabled>{element.IsEnabled}</IsEnabled>");
                    sb.AppendLine("    </Element>");
                }

                sb.AppendLine("  </Elements>");
                sb.AppendLine("</UIElementsExport>");

                return sb.ToString();
            });
        }

        private string EscapeXml(string value)
        {
            if (string.IsNullOrEmpty(value))
                return "";
            return System.Security.SecurityElement.Escape(value);
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

        #region Operation Status

        /// <summary>
        /// Set operation status in status bar and show progress panel with wait cursor
        /// </summary>
        private void SetOperationStatus(string status, string progress = "", string detail = null)
        {
            Dispatcher.Invoke(() =>
            {
                sbOperationStatus.Text = status;
                sbOperationProgress.Text = progress;

                // Show progress indicator panel
                pnlProgressIndicator.Visibility = Visibility.Visible;
                txtProgressMessage.Text = status;
                txtProgressDetail.Text = detail ?? "Lütfen bekleyiniz, UI element bilgileri toplanıyor...";
                txtProgressStep.Text = progress;
                txtProgressPercent.Text = "";
                sbProgressBar.IsIndeterminate = true;

                // Set wait cursor
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            });
        }

        /// <summary>
        /// Clear operation status and hide progress panel, restore cursor
        /// </summary>
        private void ClearOperationStatus()
        {
            Dispatcher.Invoke(() =>
            {
                sbOperationStatus.Text = "";
                sbOperationProgress.Text = "";

                // Hide progress indicator panel
                pnlProgressIndicator.Visibility = Visibility.Collapsed;
                sbProgressBar.IsIndeterminate = false;
                sbProgressBar.Value = 0;
                sbMainProgressBar.Value = 0;
                txtProgressPercent.Text = "";
                txtProgressStep.Text = "";

                // Restore cursor
                Mouse.OverrideCursor = null;
            });
        }

        /// <summary>
        /// Update progress bar value (0-100) with detailed message
        /// </summary>
        private void SetProgressValue(int value, string message = null, string step = null)
        {
            Dispatcher.Invoke(() =>
            {
                sbProgressBar.IsIndeterminate = false;
                sbProgressBar.Value = value;
                sbMainProgressBar.Value = value;
                txtProgressPercent.Text = $"%{value}";
                if (message != null)
                {
                    txtProgressMessage.Text = message;
                }
                if (step != null)
                {
                    txtProgressStep.Text = step;
                }
            });
        }

        /// <summary>
        /// Start progress with indeterminate mode and wait cursor
        /// </summary>
        private void StartProgress(string message, string detail = null)
        {
            Dispatcher.Invoke(() =>
            {
                pnlProgressIndicator.Visibility = Visibility.Visible;
                txtProgressMessage.Text = message;
                txtProgressDetail.Text = detail ?? "Lütfen bekleyiniz, UI element bilgileri toplanıyor...";
                txtProgressIcon.Text = "⏳";
                sbProgressBar.IsIndeterminate = true;
                sbOperationStatus.Text = message;
                Mouse.OverrideCursor = System.Windows.Input.Cursors.Wait;
            });
        }

        /// <summary>
        /// Stop progress and restore cursor
        /// </summary>
        private void StopProgress()
        {
            Dispatcher.Invoke(() =>
            {
                pnlProgressIndicator.Visibility = Visibility.Collapsed;
                sbProgressBar.IsIndeterminate = false;
                sbProgressBar.Value = 0;
                sbMainProgressBar.Value = 0;
                sbOperationStatus.Text = "";
                txtProgressPercent.Text = "";
                Mouse.OverrideCursor = null;
            });
        }

        #endregion

        #region Archive Tab

        /// <summary>
        /// Initialize archive tab
        /// </summary>
        private void InitializeArchiveTab()
        {
            try
            {
                if (_archiveManager != null)
                {
                    txtArchivePath.Text = _archiveManager.ArchiveBasePath;
                    RefreshArchiveList();
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Archive initialization error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        /// <summary>
        /// Refresh archive list from disk
        /// </summary>
        private void RefreshArchiveList()
        {
            try
            {
                if (_archiveManager == null) return;

                _archiveManager.Refresh();
                lvArchive.ItemsSource = null;
                lvArchive.ItemsSource = _archiveManager.Items;
            }
            catch (Exception ex)
            {
                LogToConsole($"Archive refresh error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private void ArchiveRefresh_Click(object sender, RoutedEventArgs e)
        {
            RefreshArchiveList();
            LogToConsole("Archive list refreshed");
        }

        private void ArchiveOpenFolder_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_archiveManager != null && System.IO.Directory.Exists(_archiveManager.ArchiveBasePath))
                {
                    Process.Start("explorer.exe", _archiveManager.ArchiveBasePath);
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Failed to open archive folder: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private void ArchiveList_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // Can be used to show details of selected archive item
        }

        private void ArchiveRename_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var button = sender as System.Windows.Controls.Button;
                var itemId = button?.Tag as string;
                if (string.IsNullOrEmpty(itemId)) return;

                var item = _archiveManager.GetItem(itemId);
                if (item == null) return;

                // Simple input dialog using MessageBox
                var inputDialog = new System.Windows.Window
                {
                    Title = "Rename Archive Item",
                    Width = 400,
                    Height = 150,
                    WindowStartupLocation = WindowStartupLocation.CenterOwner,
                    Owner = this,
                    ResizeMode = ResizeMode.NoResize
                };

                var panel = new StackPanel { Margin = new Thickness(10) };
                panel.Children.Add(new TextBlock { Text = "Enter new name:", Margin = new Thickness(0, 0, 0, 5) });

                var textBox = new System.Windows.Controls.TextBox { Text = item.Name, Margin = new Thickness(0, 0, 0, 10) };
                panel.Children.Add(textBox);

                var buttonPanel = new StackPanel { Orientation = System.Windows.Controls.Orientation.Horizontal, HorizontalAlignment = System.Windows.HorizontalAlignment.Right };
                var okButton = new System.Windows.Controls.Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 10, 0) };
                var cancelButton = new System.Windows.Controls.Button { Content = "Cancel", Width = 80 };

                okButton.Click += (s, args) => { inputDialog.DialogResult = true; inputDialog.Close(); };
                cancelButton.Click += (s, args) => { inputDialog.DialogResult = false; inputDialog.Close(); };

                buttonPanel.Children.Add(okButton);
                buttonPanel.Children.Add(cancelButton);
                panel.Children.Add(buttonPanel);

                inputDialog.Content = panel;

                if (inputDialog.ShowDialog() == true && !string.IsNullOrWhiteSpace(textBox.Text))
                {
                    _archiveManager.UpdateItemName(itemId, textBox.Text);
                    RefreshArchiveList();
                    LogToConsole($"Archive item renamed to: {textBox.Text}");
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Rename error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private void ArchiveCopyLinks_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var button = sender as System.Windows.Controls.Button;
                var itemId = button?.Tag as string;
                if (string.IsNullOrEmpty(itemId)) return;

                var links = _archiveManager.GetFileLinksForClipboard(itemId);
                if (!string.IsNullOrEmpty(links))
                {
                    System.Windows.Clipboard.SetText(links);
                    LogToConsole("File links copied to clipboard");
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Copy links error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private async void ArchiveCopyContent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var button = sender as System.Windows.Controls.Button;
                var itemId = button?.Tag as string;
                if (string.IsNullOrEmpty(itemId)) return;

                var content = await _archiveManager.GetAllFileContentsForClipboard(itemId);
                if (!string.IsNullOrEmpty(content))
                {
                    System.Windows.Clipboard.SetText(content);
                    LogToConsole("File contents copied to clipboard");
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Copy content error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private void ArchiveOpenItem_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var button = sender as System.Windows.Controls.Button;
                var itemId = button?.Tag as string;
                if (string.IsNullOrEmpty(itemId)) return;

                _archiveManager.OpenInExplorer(itemId);
            }
            catch (Exception ex)
            {
                LogToConsole($"Open item error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        private void ArchiveDelete_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var button = sender as System.Windows.Controls.Button;
                var itemId = button?.Tag as string;
                if (string.IsNullOrEmpty(itemId)) return;

                var item = _archiveManager.GetItem(itemId);
                if (item == null) return;

                var result = System.Windows.MessageBox.Show(
                    $"Are you sure you want to delete this archive item?\n\n{item.Name}\n\nThis will permanently delete all files in the archive folder.",
                    "Confirm Delete",
                    MessageBoxButton.YesNo,
                    MessageBoxImage.Warning);

                if (result == MessageBoxResult.Yes)
                {
                    if (_archiveManager.DeleteItem(itemId))
                    {
                        RefreshArchiveList();
                        LogToConsole($"Archive item deleted: {item.Name}");
                    }
                    else
                    {
                        LogToConsole("Failed to delete archive item", Core.Utils.LogLevel.Error);
                    }
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Delete error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        /// <summary>
        /// Copy selected archive item's full content to clipboard (for side panel button)
        /// </summary>
        private async void ArchiveSelectedCopyContent_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedItem = lvArchive.SelectedItem as ArchiveItem;
                if (selectedItem == null)
                {
                    LogToConsole("Please select an archive item first.");
                    System.Windows.MessageBox.Show("Please select an archive item from the list first.", "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                SetOperationStatus("Copying content...", "Please wait");
                var content = await _archiveManager.GetAllFileContentsForClipboard(selectedItem.Id);
                if (!string.IsNullOrEmpty(content))
                {
                    System.Windows.Clipboard.SetText(content);
                    LogToConsole($"Full content copied to clipboard ({content.Length:N0} characters)");
                    System.Windows.MessageBox.Show($"Full report copied to clipboard!\n\nSize: {content.Length:N0} characters", "Copy Successful", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                ClearOperationStatus();
            }
            catch (Exception ex)
            {
                ClearOperationStatus();
                LogToConsole($"Copy content error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        /// <summary>
        /// Copy selected archive item's file links to clipboard (for side panel button)
        /// </summary>
        private void ArchiveSelectedCopyLinks_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedItem = lvArchive.SelectedItem as ArchiveItem;
                if (selectedItem == null)
                {
                    LogToConsole("Please select an archive item first.");
                    System.Windows.MessageBox.Show("Please select an archive item from the list first.", "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                var links = _archiveManager.GetFileLinksForClipboard(selectedItem.Id);
                if (!string.IsNullOrEmpty(links))
                {
                    System.Windows.Clipboard.SetText(links);
                    LogToConsole("File links copied to clipboard");
                    System.Windows.MessageBox.Show("File links copied to clipboard!", "Copy Successful", MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                LogToConsole($"Copy links error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        /// <summary>
        /// Open selected archive item's folder in Explorer (for side panel button)
        /// </summary>
        private void ArchiveSelectedOpen_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var selectedItem = lvArchive.SelectedItem as ArchiveItem;
                if (selectedItem == null)
                {
                    LogToConsole("Please select an archive item first.");
                    System.Windows.MessageBox.Show("Please select an archive item from the list first.", "No Selection", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                _archiveManager.OpenInExplorer(selectedItem.Id);
            }
            catch (Exception ex)
            {
                LogToConsole($"Open folder error: {ex.Message}", Core.Utils.LogLevel.Error);
            }
        }

        #endregion

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);

            _logger?.LogSection("APPLICATION SHUTDOWN");
            _logger?.LogInfo("Cleaning up resources...");

            // Cleanup
            _mouseHook?.Dispose();
            _hotkeyService?.Dispose();
            _memoryTimer?.Stop();
            _mouseTimer?.Stop();
            _inspectionCts?.Cancel();
            _floatingWindow?.Close();

            _logger?.LogInfo("All resources cleaned up successfully");
            _logger?.Dispose();
        }
    }
}