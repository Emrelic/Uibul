using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace UIElementInspector.Core.Utils
{
    /// <summary>
    /// Helper class for taking screenshots of elements, regions, or full screen
    /// </summary>
    public static class ScreenshotHelper
    {
        #region Native Methods

        [DllImport("user32.dll")]
        private static extern IntPtr GetDesktopWindow();

        [DllImport("user32.dll")]
        private static extern IntPtr GetWindowDC(IntPtr hWnd);

        [DllImport("user32.dll")]
        private static extern int ReleaseDC(IntPtr hWnd, IntPtr hDC);

        [DllImport("gdi32.dll")]
        private static extern bool BitBlt(IntPtr hdcDest, int nXDest, int nYDest,
            int nWidth, int nHeight, IntPtr hdcSrc, int nXSrc, int nYSrc, uint dwRop);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleBitmap(IntPtr hdc, int nWidth, int nHeight);

        [DllImport("gdi32.dll")]
        private static extern IntPtr CreateCompatibleDC(IntPtr hdc);

        [DllImport("gdi32.dll")]
        private static extern IntPtr SelectObject(IntPtr hdc, IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteObject(IntPtr hObject);

        [DllImport("gdi32.dll")]
        private static extern bool DeleteDC(IntPtr hdc);

        private const uint SRCCOPY = 0x00CC0020;

        #endregion

        /// <summary>
        /// Captures a screenshot of the entire screen
        /// </summary>
        public static Bitmap CaptureFullScreen()
        {
            try
            {
                var screenBounds = System.Windows.Forms.Screen.PrimaryScreen.Bounds;
                return CaptureRegion(new Rectangle(0, 0, screenBounds.Width, screenBounds.Height));
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to capture full screen: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Captures a screenshot of a specific region
        /// </summary>
        public static Bitmap CaptureRegion(Rectangle region)
        {
            IntPtr desktopDC = IntPtr.Zero;
            IntPtr memoryDC = IntPtr.Zero;
            IntPtr bitmap = IntPtr.Zero;
            IntPtr oldBitmap = IntPtr.Zero;

            try
            {
                // Get device context of the desktop
                desktopDC = GetWindowDC(GetDesktopWindow());
                memoryDC = CreateCompatibleDC(desktopDC);

                // Create bitmap
                bitmap = CreateCompatibleBitmap(desktopDC, region.Width, region.Height);
                oldBitmap = SelectObject(memoryDC, bitmap);

                // Copy screen to bitmap
                BitBlt(memoryDC, 0, 0, region.Width, region.Height,
                       desktopDC, region.X, region.Y, SRCCOPY);

                // Select old bitmap back
                SelectObject(memoryDC, oldBitmap);

                // Create managed bitmap from unmanaged
                var screenshot = Image.FromHbitmap(bitmap);
                return screenshot as Bitmap;
            }
            finally
            {
                // Clean up
                if (oldBitmap != IntPtr.Zero) SelectObject(memoryDC, oldBitmap);
                if (bitmap != IntPtr.Zero) DeleteObject(bitmap);
                if (memoryDC != IntPtr.Zero) DeleteDC(memoryDC);
                if (desktopDC != IntPtr.Zero) ReleaseDC(GetDesktopWindow(), desktopDC);
            }
        }

        /// <summary>
        /// Captures a screenshot of a specific window
        /// </summary>
        public static Bitmap CaptureWindow(IntPtr windowHandle)
        {
            try
            {
                // Get window rectangle
                var rect = new System.Windows.Forms.Form();

                if (windowHandle != IntPtr.Zero)
                {
                    RECT windowRect;
                    GetWindowRect(windowHandle, out windowRect);
                    var bounds = new Rectangle(
                        windowRect.Left,
                        windowRect.Top,
                        windowRect.Right - windowRect.Left,
                        windowRect.Bottom - windowRect.Top
                    );
                    return CaptureRegion(bounds);
                }

                return null;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to capture window: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Captures a screenshot of an element based on its bounds
        /// </summary>
        public static Bitmap CaptureElement(Rect elementBounds)
        {
            try
            {
                var region = new Rectangle(
                    (int)elementBounds.X,
                    (int)elementBounds.Y,
                    (int)elementBounds.Width,
                    (int)elementBounds.Height
                );
                return CaptureRegion(region);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to capture element: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Highlights a region on screen by drawing a rectangle
        /// </summary>
        public static void HighlightRegion(Rect region, int durationMs = 2000)
        {
            try
            {
                var highlight = new System.Windows.Forms.Form
                {
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None,
                    BackColor = System.Drawing.Color.Red,
                    Opacity = 0.3,
                    TopMost = true,
                    ShowInTaskbar = false,
                    StartPosition = System.Windows.Forms.FormStartPosition.Manual,
                    Location = new System.Drawing.Point((int)region.X, (int)region.Y),
                    Size = new System.Drawing.Size((int)region.Width, (int)region.Height)
                };

                // Create border effect
                var borderForm = new System.Windows.Forms.Form
                {
                    FormBorderStyle = System.Windows.Forms.FormBorderStyle.None,
                    BackColor = System.Drawing.Color.Red,
                    TopMost = true,
                    ShowInTaskbar = false,
                    StartPosition = System.Windows.Forms.FormStartPosition.Manual,
                    TransparencyKey = System.Drawing.Color.White,
                    Location = new System.Drawing.Point((int)region.X - 2, (int)region.Y - 2),
                    Size = new System.Drawing.Size((int)region.Width + 4, (int)region.Height + 4)
                };

                // Create hollow rectangle effect
                using (var g = borderForm.CreateGraphics())
                {
                    g.Clear(System.Drawing.Color.White);
                    using (var pen = new System.Drawing.Pen(System.Drawing.Color.Red, 3))
                    {
                        g.DrawRectangle(pen, 1, 1, borderForm.Width - 3, borderForm.Height - 3);
                    }
                }

                borderForm.Show();

                // Auto-close after duration
                var timer = new System.Windows.Forms.Timer { Interval = durationMs };
                timer.Tick += (s, e) =>
                {
                    timer.Stop();
                    borderForm.Close();
                    timer.Dispose();
                };
                timer.Start();
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to highlight region: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Converts Bitmap to BitmapImage for WPF
        /// </summary>
        public static BitmapImage ConvertToBitmapImage(Bitmap bitmap)
        {
            using (var memory = new MemoryStream())
            {
                bitmap.Save(memory, ImageFormat.Png);
                memory.Position = 0;

                var bitmapImage = new BitmapImage();
                bitmapImage.BeginInit();
                bitmapImage.StreamSource = memory;
                bitmapImage.CacheOption = BitmapCacheOption.OnLoad;
                bitmapImage.EndInit();

                return bitmapImage;
            }
        }

        /// <summary>
        /// Converts Bitmap to byte array
        /// </summary>
        public static byte[] ConvertToByteArray(Bitmap bitmap)
        {
            using (var stream = new MemoryStream())
            {
                bitmap.Save(stream, ImageFormat.Png);
                return stream.ToArray();
            }
        }

        /// <summary>
        /// Saves bitmap to file
        /// </summary>
        public static void SaveToFile(Bitmap bitmap, string filePath)
        {
            try
            {
                var extension = Path.GetExtension(filePath)?.ToLower();
                ImageFormat format = ImageFormat.Png;

                switch (extension)
                {
                    case ".jpg":
                    case ".jpeg":
                        format = ImageFormat.Jpeg;
                        break;
                    case ".bmp":
                        format = ImageFormat.Bmp;
                        break;
                    case ".gif":
                        format = ImageFormat.Gif;
                        break;
                    case ".tiff":
                        format = ImageFormat.Tiff;
                        break;
                }

                bitmap.Save(filePath, format);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to save screenshot: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Captures with mouse cursor
        /// </summary>
        public static Bitmap CaptureWithCursor(Rectangle region)
        {
            var screenshot = CaptureRegion(region);

            // Draw cursor on screenshot
            try
            {
                using (var g = Graphics.FromImage(screenshot))
                {
                    var cursorPosition = System.Windows.Forms.Cursor.Position;
                    cursorPosition.X -= region.X;
                    cursorPosition.Y -= region.Y;

                    // Draw simple cursor representation
                    using (var pen = new System.Drawing.Pen(System.Drawing.Color.Black, 2))
                    {
                        g.DrawLine(pen, cursorPosition.X - 10, cursorPosition.Y, cursorPosition.X + 10, cursorPosition.Y);
                        g.DrawLine(pen, cursorPosition.X, cursorPosition.Y - 10, cursorPosition.X, cursorPosition.Y + 10);
                    }
                }
            }
            catch { }

            return screenshot;
        }

        /// <summary>
        /// Creates a thumbnail from bitmap
        /// </summary>
        public static Bitmap CreateThumbnail(Bitmap original, int maxWidth, int maxHeight)
        {
            var ratioX = (double)maxWidth / original.Width;
            var ratioY = (double)maxHeight / original.Height;
            var ratio = Math.Min(ratioX, ratioY);

            var newWidth = (int)(original.Width * ratio);
            var newHeight = (int)(original.Height * ratio);

            var thumbnail = new Bitmap(newWidth, newHeight);

            using (var g = Graphics.FromImage(thumbnail))
            {
                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                g.DrawImage(original, 0, 0, newWidth, newHeight);
            }

            return thumbnail;
        }

        #region Native Structures

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [DllImport("user32.dll")]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        #endregion
    }
}