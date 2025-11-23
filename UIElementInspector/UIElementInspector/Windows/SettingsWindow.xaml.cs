using System;
using System.IO;
using System.Windows;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Windows
{
    /// <summary>
    /// Settings window for configuring application preferences
    /// </summary>
    public partial class SettingsWindow : Window
    {
        private AppSettings _settings;
        private bool _settingsChanged = false;

        public AppSettings Settings => _settings;
        public bool SettingsChanged => _settingsChanged;

        public SettingsWindow()
        {
            InitializeComponent();
            _settings = AppSettings.Load();
            LoadSettings();
        }

        public SettingsWindow(AppSettings existingSettings)
        {
            InitializeComponent();
            _settings = existingSettings ?? AppSettings.Load();
            LoadSettings();
        }

        private void LoadSettings()
        {
            // Collection Profile
            switch (_settings.DefaultCollectionProfile)
            {
                case CollectionProfile.Quick:
                    rbQuick.IsChecked = true;
                    break;
                case CollectionProfile.Standard:
                    rbStandard.IsChecked = true;
                    break;
                case CollectionProfile.Full:
                    rbFull.IsChecked = true;
                    break;
                case CollectionProfile.Custom:
                    rbCustom.IsChecked = true;
                    break;
            }

            // Export Settings
            chkExportCsv.IsChecked = _settings.ExportFormats.Contains("CSV");
            chkExportJson.IsChecked = _settings.ExportFormats.Contains("JSON");
            chkExportXml.IsChecked = _settings.ExportFormats.Contains("XML");
            chkExportHtml.IsChecked = _settings.ExportFormats.Contains("HTML");

            txtExportPath.Text = _settings.ExportDirectory;
            chkAutoExport.IsChecked = _settings.AutoExportOnCapture;
            chkOrganizeByDate.IsChecked = _settings.OrganizeExportsByDate;

            // Screenshot Settings
            cmbScreenshotFormat.SelectedIndex = _settings.ScreenshotFormat switch
            {
                "PNG" => 0,
                "JPEG" => 1,
                "BMP" => 2,
                _ => 0
            };

            sliderJpegQuality.Value = _settings.JpegQuality;
            chkAutoScreenshot.IsChecked = _settings.AutoCaptureScreenshot;
            chkIncludeTimestamp.IsChecked = _settings.IncludeTimestampInFilename;

            // Performance Settings
            sliderHoverDelay.Value = _settings.MouseHoverDelay;
            sliderMaxTreeDepth.Value = _settings.MaxTreeDepth;
            chkEnableThrottling.IsChecked = _settings.EnableDetectionThrottling;
            chkCacheElements.IsChecked = _settings.CacheElements;

            // UI Settings
            chkAlwaysOnTop.IsChecked = _settings.AlwaysOnTop;
            chkMinimizeToTray.IsChecked = _settings.MinimizeToTray;
            chkShowNotifications.IsChecked = _settings.ShowNotifications;
            chkShowTooltips.IsChecked = _settings.ShowTooltips;
        }

        private void SaveSettings()
        {
            // Collection Profile
            if (rbQuick.IsChecked == true)
                _settings.DefaultCollectionProfile = CollectionProfile.Quick;
            else if (rbStandard.IsChecked == true)
                _settings.DefaultCollectionProfile = CollectionProfile.Standard;
            else if (rbFull.IsChecked == true)
                _settings.DefaultCollectionProfile = CollectionProfile.Full;
            else if (rbCustom.IsChecked == true)
                _settings.DefaultCollectionProfile = CollectionProfile.Custom;

            // Export Settings
            _settings.ExportFormats.Clear();
            if (chkExportCsv.IsChecked == true) _settings.ExportFormats.Add("CSV");
            if (chkExportJson.IsChecked == true) _settings.ExportFormats.Add("JSON");
            if (chkExportXml.IsChecked == true) _settings.ExportFormats.Add("XML");
            if (chkExportHtml.IsChecked == true) _settings.ExportFormats.Add("HTML");

            _settings.ExportDirectory = txtExportPath.Text;
            _settings.AutoExportOnCapture = chkAutoExport.IsChecked == true;
            _settings.OrganizeExportsByDate = chkOrganizeByDate.IsChecked == true;

            // Screenshot Settings
            _settings.ScreenshotFormat = cmbScreenshotFormat.SelectedIndex switch
            {
                0 => "PNG",
                1 => "JPEG",
                2 => "BMP",
                _ => "PNG"
            };

            _settings.JpegQuality = (int)sliderJpegQuality.Value;
            _settings.AutoCaptureScreenshot = chkAutoScreenshot.IsChecked == true;
            _settings.IncludeTimestampInFilename = chkIncludeTimestamp.IsChecked == true;

            // Performance Settings
            _settings.MouseHoverDelay = (int)sliderHoverDelay.Value;
            _settings.MaxTreeDepth = (int)sliderMaxTreeDepth.Value;
            _settings.EnableDetectionThrottling = chkEnableThrottling.IsChecked == true;
            _settings.CacheElements = chkCacheElements.IsChecked == true;

            // UI Settings
            _settings.AlwaysOnTop = chkAlwaysOnTop.IsChecked == true;
            _settings.MinimizeToTray = chkMinimizeToTray.IsChecked == true;
            _settings.ShowNotifications = chkShowNotifications.IsChecked == true;
            _settings.ShowTooltips = chkShowTooltips.IsChecked == true;

            // Save to disk
            _settings.Save();
            _settingsChanged = true;
        }

        private void BrowseExportPath_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog
            {
                Description = "Select export directory",
                SelectedPath = txtExportPath.Text
            };

            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                txtExportPath.Text = dialog.SelectedPath;
            }
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // Validate settings
            if (string.IsNullOrWhiteSpace(txtExportPath.Text))
            {
                System.Windows.MessageBox.Show("Please specify an export directory.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!_settings.ExportFormats.Any())
            {
                System.Windows.MessageBox.Show("Please select at least one export format.", "Validation Error",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Create export directory if it doesn't exist
            try
            {
                if (!Directory.Exists(txtExportPath.Text))
                {
                    Directory.CreateDirectory(txtExportPath.Text);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Error creating export directory: {ex.Message}", "Error",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SaveSettings();
            DialogResult = true;
            Close();
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void ResetDefaults_Click(object sender, RoutedEventArgs e)
        {
            var result = System.Windows.MessageBox.Show(
                "Are you sure you want to reset all settings to their default values?",
                "Reset Settings",
                MessageBoxButton.YesNo,
                MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _settings = AppSettings.CreateDefault();
                LoadSettings();
            }
        }
    }
}
