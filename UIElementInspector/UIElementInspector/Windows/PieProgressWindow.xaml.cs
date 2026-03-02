using System;
using System.Windows;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Threading.Tasks;

using WpfPoint = System.Windows.Point;
using WpfSize = System.Windows.Size;
using WpfColor = System.Windows.Media.Color;

namespace UIElementInspector.Windows
{
    /// <summary>
    /// Floating pie-chart progress overlay that appears on top of all windows
    /// during capture operations. Shows a circular progress indicator that fills
    /// clockwise from 0 to 360 degrees as the capture progresses.
    /// </summary>
    public partial class PieProgressWindow : Window
    {
        private const double CircleRadius = 25.0;
        private readonly WpfPoint _center = new WpfPoint(25, 25);
        private System.Windows.Threading.DispatcherTimer _timeoutTimer;

        public PieProgressWindow()
        {
            InitializeComponent();
            _timeoutTimer = new System.Windows.Threading.DispatcherTimer();
            _timeoutTimer.Interval = TimeSpan.FromSeconds(60);
            _timeoutTimer.Tick += (s, e) =>
            {
                _timeoutTimer.Stop();
                this.Hide();
            };
        }

        /// <summary>
        /// Show the overlay near the specified screen position
        /// </summary>
        public void ShowAtPosition(double screenX, double screenY)
        {
            // Pin to bottom-right corner of screen
            var screenWidth = SystemParameters.PrimaryScreenWidth;
            var screenHeight = SystemParameters.PrimaryScreenHeight;

            this.Left = screenWidth - this.Width - 16;
            this.Top = screenHeight - this.Height - 50; // above taskbar

            // Reset state
            txtPercent.Visibility = Visibility.Visible;
            txtCheckmark.Visibility = Visibility.Collapsed;
            txtCheckmark.Text = "\u2714"; // Reset to checkmark
            txtCheckmark.Foreground = new SolidColorBrush(WpfColor.FromRgb(76, 175, 80));
            txtPercent.Text = "0%";
            txtStep.Text = "";
            txtMessage.Text = "";
            pieArc.Fill = new SolidColorBrush(WpfColor.FromRgb(25, 118, 210)); // #1976D2
            this.Opacity = 1.0;

            UpdatePieGeometry(0);

            _timeoutTimer.Stop();
            _timeoutTimer.Start();
            this.Show();
        }

        /// <summary>
        /// Update progress: percentage (0-100), optional message and step text
        /// </summary>
        public void UpdateProgress(int percentage, string? message = null, string? step = null)
        {
            Dispatcher.Invoke(() =>
            {
                int clamped = Math.Max(0, Math.Min(100, percentage));
                txtPercent.Text = $"{clamped}%";

                if (message != null)
                    txtMessage.Text = message;
                if (step != null)
                    txtStep.Text = step;

                UpdatePieGeometry(clamped);
            });
        }

        /// <summary>
        /// Show green completion state with checkmark
        /// </summary>
        public async void ShowCompleted()
        {
            _timeoutTimer.Stop();
            Dispatcher.Invoke(() =>
            {
                UpdatePieGeometry(100);
                pieArc.Fill = new SolidColorBrush(WpfColor.FromRgb(76, 175, 80)); // #4CAF50 green
                txtPercent.Visibility = Visibility.Collapsed;
                txtCheckmark.Text = "\u2714";
                txtCheckmark.Foreground = new SolidColorBrush(WpfColor.FromRgb(76, 175, 80));
                txtCheckmark.Visibility = Visibility.Visible;
                txtStep.Text = "Tamamlandi!";
                txtMessage.Text = "";
            });

            // Keep visible briefly then fade out
            await Task.Delay(1200);

            // Fade out animation
            for (double opacity = 1.0; opacity >= 0; opacity -= 0.1)
            {
                Dispatcher.Invoke(() => this.Opacity = opacity);
                await Task.Delay(30);
            }

            Dispatcher.Invoke(() => this.Hide());
        }

        /// <summary>
        /// Show red error state
        /// </summary>
        public async void ShowError()
        {
            _timeoutTimer.Stop();
            Dispatcher.Invoke(() =>
            {
                pieArc.Fill = new SolidColorBrush(WpfColor.FromRgb(211, 47, 47)); // #D32F2F red
                txtPercent.Visibility = Visibility.Collapsed;
                txtCheckmark.Text = "\u2718"; // X mark
                txtCheckmark.Foreground = new SolidColorBrush(WpfColor.FromRgb(211, 47, 47));
                txtCheckmark.Visibility = Visibility.Visible;
                txtStep.Text = "Hata!";
                txtMessage.Text = "";
            });

            await Task.Delay(2000);
            Dispatcher.Invoke(() => this.Hide());
        }

        /// <summary>
        /// Close and hide the overlay
        /// </summary>
        public void CloseOverlay()
        {
            Dispatcher.Invoke(() =>
            {
                this.Hide();
            });
        }

        /// <summary>
        /// Draw the pie arc geometry for the given percentage (0-100)
        /// </summary>
        private void UpdatePieGeometry(int percentage)
        {
            if (percentage <= 0)
            {
                pieArc.Data = Geometry.Empty;
                return;
            }

            if (percentage >= 100)
            {
                // Full circle - use EllipseGeometry
                pieArc.Data = new EllipseGeometry(_center, CircleRadius, CircleRadius);
                return;
            }

            // Calculate arc angle (0-360 degrees)
            double angle = (percentage / 100.0) * 360.0;

            // Start from 12 o'clock (top center), go clockwise
            var startPoint = new WpfPoint(_center.X, _center.Y - CircleRadius);
            var endPoint = CalculateArcPoint(angle);

            bool isLargeArc = angle > 180.0;

            var pathFigure = new PathFigure
            {
                StartPoint = _center,
                IsClosed = true
            };

            // Line from center to start (12 o'clock)
            pathFigure.Segments.Add(new LineSegment(startPoint, true));

            // Arc from start to end point
            pathFigure.Segments.Add(new ArcSegment
            {
                Point = endPoint,
                Size = new WpfSize(CircleRadius, CircleRadius),
                SweepDirection = SweepDirection.Clockwise,
                IsLargeArc = isLargeArc
            });

            // Line back to center (auto-closed)

            var pathGeometry = new PathGeometry();
            pathGeometry.Figures.Add(pathFigure);

            pieArc.Data = pathGeometry;
        }

        /// <summary>
        /// Calculate point on circle for given angle (degrees, 0 = 12 o'clock, clockwise)
        /// </summary>
        private WpfPoint CalculateArcPoint(double angleDegrees)
        {
            // Convert to radians, offset by -90 so 0 degrees = 12 o'clock
            double angleRadians = (angleDegrees - 90.0) * (Math.PI / 180.0);
            double x = _center.X + CircleRadius * Math.Cos(angleRadians);
            double y = _center.Y + CircleRadius * Math.Sin(angleRadians);
            return new WpfPoint(x, y);
        }

        /// <summary>
        /// Prevent closing - just hide instead
        /// </summary>
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
            base.OnClosing(e);
        }
    }
}
