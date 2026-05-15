using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using Newtonsoft.Json;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Manages the archive of captured UI elements
    /// </summary>
    public class ArchiveManager
    {
        private readonly string _archiveBasePath;
        private readonly string _indexFilePath;
        private ArchiveIndex _index;
        private readonly object _lockObject = new object();

        public event EventHandler<ArchiveItem> ItemAdded;
        public event EventHandler<string> ItemDeleted;
        public event EventHandler<ArchiveItem> ItemUpdated;
        public event EventHandler IndexLoaded;

        public string ArchiveBasePath => _archiveBasePath;
        public IReadOnlyList<ArchiveItem> Items => _index?.Items?.AsReadOnly() ?? new List<ArchiveItem>().AsReadOnly();

        public ArchiveManager()
        {
            // Archive folder in the application's data directory
            var appDataPath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "UIElementInspector",
                "Archive");

            _archiveBasePath = appDataPath;
            _indexFilePath = Path.Combine(_archiveBasePath, "archive_index.json");

            EnsureArchiveDirectoryExists();
            LoadIndex();
        }

        public ArchiveManager(string customBasePath)
        {
            _archiveBasePath = customBasePath;
            _indexFilePath = Path.Combine(_archiveBasePath, "archive_index.json");

            EnsureArchiveDirectoryExists();
            LoadIndex();
        }

        private void EnsureArchiveDirectoryExists()
        {
            if (!Directory.Exists(_archiveBasePath))
            {
                Directory.CreateDirectory(_archiveBasePath);
            }
        }

        private void LoadIndex()
        {
            try
            {
                if (File.Exists(_indexFilePath))
                {
                    var json = File.ReadAllText(_indexFilePath);
                    _index = JsonConvert.DeserializeObject<ArchiveIndex>(json) ?? new ArchiveIndex();

                    // Validate and clean up missing folders
                    _index.Items.RemoveAll(item => !Directory.Exists(item.FolderPath));
                }
                else
                {
                    _index = new ArchiveIndex();
                }

                IndexLoaded?.Invoke(this, EventArgs.Empty);
            }
            catch (Exception)
            {
                _index = new ArchiveIndex();
            }
        }

        public void SaveIndex()
        {
            lock (_lockObject)
            {
                try
                {
                    _index.LastUpdated = DateTime.Now;
                    var json = JsonConvert.SerializeObject(_index, Formatting.Indented);
                    File.WriteAllText(_indexFilePath, json);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Failed to save archive index: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// Create a new archive item and return its folder path
        /// </summary>
        public ArchiveItem CreateArchiveItem(string name = null, string captureType = "FullCapture")
        {
            var timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
            var folderName = $"Capture_{timestamp}";
            var folderPath = Path.Combine(_archiveBasePath, folderName);

            Directory.CreateDirectory(folderPath);

            var item = new ArchiveItem
            {
                Id = Guid.NewGuid().ToString(),
                Name = name ?? $"Capture {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
                FolderPath = folderPath,
                CaptureTime = DateTime.Now,
                CaptureType = captureType
            };

            lock (_lockObject)
            {
                _index.Items.Insert(0, item); // Add to beginning (newest first)
            }

            SaveIndex();
            ItemAdded?.Invoke(this, item);

            return item;
        }

        /// <summary>
        /// Add files to an archive item
        /// </summary>
        public void AddFilesToItem(ArchiveItem item, List<string> filePaths)
        {
            item.FilePaths.AddRange(filePaths);
            item.FileCount = item.FilePaths.Count;
            SaveIndex();
            ItemUpdated?.Invoke(this, item);
        }

        /// <summary>
        /// Update archive item name
        /// </summary>
        public void UpdateItemName(string itemId, string newName)
        {
            var item = _index.Items.FirstOrDefault(i => i.Id == itemId);
            if (item != null)
            {
                item.Name = newName;
                SaveIndex();
                ItemUpdated?.Invoke(this, item);
            }
        }

        /// <summary>
        /// Update archive item notes
        /// </summary>
        public void UpdateItemNotes(string itemId, string notes)
        {
            var item = _index.Items.FirstOrDefault(i => i.Id == itemId);
            if (item != null)
            {
                item.Notes = notes;
                SaveIndex();
                ItemUpdated?.Invoke(this, item);
            }
        }

        /// <summary>
        /// Delete an archive item and its folder
        /// </summary>
        public bool DeleteItem(string itemId)
        {
            var item = _index.Items.FirstOrDefault(i => i.Id == itemId);
            if (item == null) return false;

            try
            {
                // Delete the folder
                if (Directory.Exists(item.FolderPath))
                {
                    Directory.Delete(item.FolderPath, true);
                }

                lock (_lockObject)
                {
                    _index.Items.Remove(item);
                }

                SaveIndex();
                ItemDeleted?.Invoke(this, itemId);
                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Failed to delete archive item: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Get file paths for clipboard
        /// </summary>
        public string GetFileLinksForClipboard(string itemId)
        {
            var item = _index.Items.FirstOrDefault(i => i.Id == itemId);
            if (item == null) return string.Empty;

            var lines = new List<string>
            {
                $"Archive Item: {item.Name}",
                $"Capture Time: {item.CaptureTime:yyyy-MM-dd HH:mm:ss}",
                $"Folder: {item.FolderPath}",
                "",
                "Files:"
            };

            foreach (var filePath in item.FilePaths)
            {
                lines.Add(filePath);
            }

            return string.Join(Environment.NewLine, lines);
        }

        /// <summary>
        /// Get all file contents concatenated for clipboard
        /// </summary>
        public async Task<string> GetAllFileContentsForClipboard(string itemId)
        {
            var item = _index.Items.FirstOrDefault(i => i.Id == itemId);
            if (item == null) return string.Empty;

            var sb = new System.Text.StringBuilder();
            sb.AppendLine("================================================================================");
            sb.AppendLine($"ARCHIVE ITEM: {item.Name}");
            sb.AppendLine($"Capture Time: {item.CaptureTime:yyyy-MM-dd HH:mm:ss}");
            sb.AppendLine($"Folder: {item.FolderPath}");
            sb.AppendLine("================================================================================");
            sb.AppendLine();

            foreach (var filePath in item.FilePaths)
            {
                if (!File.Exists(filePath)) continue;

                var fileName = Path.GetFileName(filePath);
                var extension = Path.GetExtension(filePath).ToLowerInvariant();

                // Skip binary files
                if (extension == ".png" || extension == ".jpg" || extension == ".jpeg" ||
                    extension == ".gif" || extension == ".bmp" || extension == ".ico")
                {
                    sb.AppendLine($"--- {fileName} ---");
                    sb.AppendLine("[Binary image file - path only]");
                    sb.AppendLine(filePath);
                    sb.AppendLine();
                    continue;
                }

                try
                {
                    var content = await File.ReadAllTextAsync(filePath);
                    sb.AppendLine($"--- {fileName} ---");
                    sb.AppendLine(content);
                    sb.AppendLine();
                }
                catch (Exception ex)
                {
                    sb.AppendLine($"--- {fileName} ---");
                    sb.AppendLine($"[Error reading file: {ex.Message}]");
                    sb.AppendLine();
                }
            }

            return sb.ToString();
        }

        /// <summary>
        /// Copy folder to desktop
        /// </summary>
        public string CopyToDesktop(string itemId)
        {
            var item = _index.Items.FirstOrDefault(i => i.Id == itemId);
            if (item == null || !Directory.Exists(item.FolderPath)) return null;

            var desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            var destFolderName = Path.GetFileName(item.FolderPath);
            var destPath = Path.Combine(desktopPath, destFolderName);

            // Handle duplicate folder names
            int counter = 1;
            while (Directory.Exists(destPath))
            {
                destPath = Path.Combine(desktopPath, $"{destFolderName}_{counter}");
                counter++;
            }

            CopyDirectory(item.FolderPath, destPath);
            return destPath;
        }

        private void CopyDirectory(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);

            foreach (var file in Directory.GetFiles(sourceDir))
            {
                var destFile = Path.Combine(destDir, Path.GetFileName(file));
                File.Copy(file, destFile, true);
            }

            foreach (var subDir in Directory.GetDirectories(sourceDir))
            {
                var destSubDir = Path.Combine(destDir, Path.GetFileName(subDir));
                CopyDirectory(subDir, destSubDir);
            }
        }

        /// <summary>
        /// Get archive item by ID
        /// </summary>
        public ArchiveItem GetItem(string itemId)
        {
            return _index.Items.FirstOrDefault(i => i.Id == itemId);
        }

        /// <summary>
        /// Refresh the index from disk
        /// </summary>
        public void Refresh()
        {
            LoadIndex();
        }

        /// <summary>
        /// Open folder in explorer
        /// </summary>
        public void OpenInExplorer(string itemId)
        {
            var item = _index.Items.FirstOrDefault(i => i.Id == itemId);
            if (item != null && Directory.Exists(item.FolderPath))
            {
                System.Diagnostics.Process.Start("explorer.exe", item.FolderPath);
            }
        }
    }
}
