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

        // ── Tarih Atlası karesi (F11) ────────────────────────────────────────
        // Kayıt klasörü KASTEN OneDrive DIŞINDA: Desktop ve Belgeler bu
        // makinede OneDrive'a bağlı (ölçüldü), her kare buluta yüklenirdi.
        public string AtlasKlasoru { get; set; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "TarihAtlasiKare");

        // Kısayol. ⚠️ ÖLÇÜLDÜ: F11 işletim sisteminde boştur ama global olarak
        // kaydedilince Chrome F11'i ARTIK GÖRMEZ — tarayıcının tam ekranı
        // araç açıkken çalışmaz. Bu bedeli istemiyorsanız "Ctrl+F11" yazın;
        // o kombinasyon da ölçüldü, boş ve çakışmasız.
        public string AtlasKisayolu { get; set; } = "F11";

        // En uzun kenar tavanı. Token maliyeti YALNIZ piksel sayısına bağlı
        // olduğu için tek gerçek tasarruf budur (format/kalite hiçbir şey
        // değiştirmez). 0 = küçültme yok. Küçük kareler ASLA büyütülmez.
        public int AtlasEnUzunKenar { get; set; } = 1200;

        // Klasörde tutulacak kare sayısı; eskiler otomatik silinir.
        public int AtlasSonKareSayisi { get; set; } = 50;

        // Başlıkta damga yoksa kare yine de alınsın mı? true ise alınır ve
        // şerit kırmızı zeminle "TARİH/KOORDİNAT OKUNAMADI" der (tarih
        // uydurulmaz). false ise kare hiç alınmaz.
        // ⚠️ Başta false idi; atlas sayfası damgayı henüz yazmadığı için
        // kısayol hiç çalışmıyor göründü. Varsayılan true olmalı.
        public bool AtlasDamgasizIzin { get; set; } = true;

        // Performance Settings
        public int MouseHoverDelay { get; set; } = 500; // milliseconds
        public int MaxTreeDepth { get; set; } = 20;
        public bool EnableDetectionThrottling { get; set; } = true;
        public bool CacheElements { get; set; } = true;

        // UI Settings
        public bool AlwaysOnTop { get; set; } = false;
        public bool ShowNotifications { get; set; } = true;
        public bool ShowTooltips { get; set; } = true;

        // ── Otomatik güncelleme ──────────────────────────────────────────────
        // Uygulama açılışta GitHub Releases'e bakar. Kontrol ARKA PLANDA ve
        // sessizdir: güncelleme yoksa hiçbir şey gösterilmez, ağ yoksa da
        // kullanıcı hata görmez (yalnız konsola yazılır). Yalnızca gerçekten
        // yeni sürüm varsa bildirim çıkar.
        public bool OtomatikGuncellemeKontrolu { get; set; } = true;

        // Aynı gün içinde tekrar tekrar sorulmasın diye son kontrol zamanı.
        public DateTime SonGuncellemeKontrolu { get; set; } = DateTime.MinValue;

        // Kaç saatte bir kontrol edilsin.
        public int GuncellemeKontrolAraligiSaat { get; set; } = 24;

        // "sahip/depo" biçiminde GitHub deposu. Çatallayan biri burayı değiştirir.
        public string GuncellemeDeposu { get; set; } = "Emrelic/Uibul";

        // Kullanıcının "bu sürümü atla" dediği etiket.
        public string AtlananSurum { get; set; } = "";

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
