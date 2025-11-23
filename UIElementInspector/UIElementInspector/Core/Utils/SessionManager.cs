using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using Newtonsoft.Json;
using UIElementInspector.Core.Models;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Manages inspection session import/export operations
    /// </summary>
    public class SessionManager
    {
        private static readonly string DefaultSessionsDirectory = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "UI Inspector Sessions");

        /// <summary>
        /// Exports a session to a file
        /// </summary>
        public static bool ExportSession(InspectionSession session, string filePath, bool compress = true)
        {
            try
            {
                // Validate session
                var errors = session.Validate();
                if (errors.Any())
                {
                    throw new InvalidOperationException($"Session validation failed: {string.Join(", ", errors)}");
                }

                // Update last modified date
                session.LastModifiedDate = DateTime.Now;

                // Serialize session to JSON
                var json = JsonConvert.SerializeObject(session, Formatting.Indented, new JsonSerializerSettings
                {
                    ReferenceLoopHandling = ReferenceLoopHandling.Ignore,
                    NullValueHandling = NullValueHandling.Include
                });

                // Ensure directory exists
                var directory = Path.GetDirectoryName(filePath);
                if (!string.IsNullOrEmpty(directory) && !Directory.Exists(directory))
                {
                    Directory.CreateDirectory(directory);
                }

                if (compress)
                {
                    // Save as compressed .uisession file
                    using (var fileStream = File.Create(filePath))
                    using (var gzipStream = new GZipStream(fileStream, CompressionMode.Compress))
                    using (var writer = new StreamWriter(gzipStream))
                    {
                        writer.Write(json);
                    }
                }
                else
                {
                    // Save as plain JSON
                    File.WriteAllText(filePath, json);
                }

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error exporting session: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Imports a session from a file
        /// </summary>
        public static InspectionSession ImportSession(string filePath)
        {
            try
            {
                if (!File.Exists(filePath))
                {
                    throw new FileNotFoundException($"Session file not found: {filePath}");
                }

                string json;
                var extension = Path.GetExtension(filePath).ToLower();

                if (extension == ".uisession")
                {
                    // Read compressed file
                    using (var fileStream = File.OpenRead(filePath))
                    using (var gzipStream = new GZipStream(fileStream, CompressionMode.Decompress))
                    using (var reader = new StreamReader(gzipStream))
                    {
                        json = reader.ReadToEnd();
                    }
                }
                else
                {
                    // Read plain JSON
                    json = File.ReadAllText(filePath);
                }

                var session = JsonConvert.DeserializeObject<InspectionSession>(json);
                if (session == null)
                {
                    throw new InvalidOperationException("Failed to deserialize session");
                }

                // Validate imported session
                var errors = session.Validate();
                if (errors.Any())
                {
                    throw new InvalidOperationException($"Imported session validation failed: {string.Join(", ", errors)}");
                }

                return session;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error importing session: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Gets all saved sessions in the default directory
        /// </summary>
        public static List<SessionInfo> GetSavedSessions()
        {
            var sessions = new List<SessionInfo>();

            try
            {
                if (!Directory.Exists(DefaultSessionsDirectory))
                {
                    Directory.CreateDirectory(DefaultSessionsDirectory);
                    return sessions;
                }

                var files = Directory.GetFiles(DefaultSessionsDirectory, "*.uisession")
                    .Concat(Directory.GetFiles(DefaultSessionsDirectory, "*.json"));

                foreach (var file in files)
                {
                    try
                    {
                        var session = ImportSession(file);
                        sessions.Add(new SessionInfo
                        {
                            FilePath = file,
                            SessionId = session.SessionId,
                            SessionName = session.SessionName,
                            CreatedDate = session.CreatedDate,
                            LastModifiedDate = session.LastModifiedDate,
                            ElementCount = session.TotalElementsCollected,
                            FileSize = new FileInfo(file).Length
                        });
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error loading session {file}: {ex.Message}");
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error getting saved sessions: {ex.Message}");
            }

            return sessions.OrderByDescending(s => s.LastModifiedDate).ToList();
        }

        /// <summary>
        /// Deletes a saved session file
        /// </summary>
        public static bool DeleteSession(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error deleting session: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Creates a new session from current collected data
        /// </summary>
        public static InspectionSession CreateSession(
            string sessionName,
            List<ElementInfo> elements,
            CollectionProfile profile,
            string detectionMethod)
        {
            var session = new InspectionSession
            {
                SessionName = sessionName,
                CollectedElements = elements ?? new List<ElementInfo>(),
                CollectionProfile = profile,
                DetectionMethod = detectionMethod,
                Description = $"Session created on {DateTime.Now:yyyy-MM-dd HH:mm:ss}"
            };

            // Calculate total collection time
            if (elements != null && elements.Any())
            {
                session.TotalCollectionTime = TimeSpan.FromMilliseconds(
                    elements.Sum(e => e.CollectionDuration.TotalMilliseconds));
            }

            return session;
        }

        /// <summary>
        /// Merges multiple sessions into one
        /// </summary>
        public static InspectionSession MergeSessions(List<InspectionSession> sessions, string newSessionName)
        {
            var mergedSession = new InspectionSession
            {
                SessionName = newSessionName,
                Description = $"Merged from {sessions.Count} sessions on {DateTime.Now:yyyy-MM-dd HH:mm:ss}"
            };

            foreach (var session in sessions)
            {
                if (session.CollectedElements != null)
                    mergedSession.CollectedElements.AddRange(session.CollectedElements);

                if (session.Screenshots != null)
                {
                    foreach (var screenshot in session.Screenshots)
                    {
                        if (!mergedSession.Screenshots.ContainsKey(screenshot.Key))
                            mergedSession.Screenshots.Add(screenshot.Key, screenshot.Value);
                    }
                }

                if (session.SourceCodes != null)
                {
                    foreach (var sourceCode in session.SourceCodes)
                    {
                        if (!mergedSession.SourceCodes.ContainsKey(sourceCode.Key))
                            mergedSession.SourceCodes.Add(sourceCode.Key, sourceCode.Value);
                    }
                }

                if (session.Tags != null)
                {
                    foreach (var tag in session.Tags)
                    {
                        if (!mergedSession.Tags.Contains(tag))
                            mergedSession.Tags.Add(tag);
                    }
                }
            }

            return mergedSession;
        }

        /// <summary>
        /// Gets the default sessions directory
        /// </summary>
        public static string GetDefaultSessionsDirectory()
        {
            if (!Directory.Exists(DefaultSessionsDirectory))
            {
                Directory.CreateDirectory(DefaultSessionsDirectory);
            }
            return DefaultSessionsDirectory;
        }
    }

    /// <summary>
    /// Lightweight session information for listing
    /// </summary>
    public class SessionInfo
    {
        public string FilePath { get; set; } = string.Empty;
        public Guid SessionId { get; set; }
        public string SessionName { get; set; } = string.Empty;
        public DateTime CreatedDate { get; set; }
        public DateTime LastModifiedDate { get; set; }
        public int ElementCount { get; set; }
        public long FileSize { get; set; }

        public string FileSizeFormatted
        {
            get
            {
                string[] sizes = { "B", "KB", "MB", "GB" };
                double len = FileSize;
                int order = 0;
                while (len >= 1024 && order < sizes.Length - 1)
                {
                    order++;
                    len = len / 1024;
                }
                return $"{len:0.##} {sizes[order]}";
            }
        }
    }
}
