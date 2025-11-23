using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Shapes;
using KeyEventArgs = System.Windows.Input.KeyEventArgs;
using MouseEventArgs = System.Windows.Input.MouseEventArgs;
using MouseButtonEventArgs = System.Windows.Input.MouseButtonEventArgs;

namespace UIElementInspector.Windows
{
    /// <summary>
    /// Window for selecting a region on screen by dragging a rectangle
    /// </summary>
    public partial class RegionSelectorWindow : Window
    {
        private System.Windows.Point _startPoint;
        private bool _isSelecting;
        private Rect _selectedRegion;

        public Rect SelectedRegion
        {
            get { return _selectedRegion; }
        }

        public bool SelectionCancelled { get; private set; }

        public RegionSelectorWindow()
        {
            InitializeComponent();
            SelectionCancelled = true; // Default to cancelled
        }

        protected override void OnMouseLeftButtonDown(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonDown(e);

            // Start selection
            _startPoint = e.GetPosition(this);
            _isSelecting = true;

            // Capture mouse
            Mouse.Capture(this);

            // Show selection rectangle
            SelectionRectangle.Visibility = Visibility.Visible;
            CoordinatesDisplay.Visibility = Visibility.Visible;

            // Position the selection rectangle
            Canvas.SetLeft(SelectionRectangle, _startPoint.X);
            Canvas.SetTop(SelectionRectangle, _startPoint.Y);
            SelectionRectangle.Width = 0;
            SelectionRectangle.Height = 0;
        }

        protected override void OnMouseMove(MouseEventArgs e)
        {
            base.OnMouseMove(e);

            if (_isSelecting)
            {
                var currentPoint = e.GetPosition(this);

                // Calculate rectangle dimensions
                var x = Math.Min(currentPoint.X, _startPoint.X);
                var y = Math.Min(currentPoint.Y, _startPoint.Y);
                var width = Math.Abs(currentPoint.X - _startPoint.X);
                var height = Math.Abs(currentPoint.Y - _startPoint.Y);

                // Update selection rectangle
                Canvas.SetLeft(SelectionRectangle, x);
                Canvas.SetTop(SelectionRectangle, y);
                SelectionRectangle.Width = width;
                SelectionRectangle.Height = height;

                // Update coordinates display
                UpdateCoordinatesDisplay(x, y, width, height);

                // Store the selected region
                _selectedRegion = new Rect(x, y, width, height);
            }
            else
            {
                // Show coordinates at mouse position when not selecting
                var point = e.GetPosition(this);
                Canvas.SetLeft(CoordinatesDisplay, point.X + 10);
                Canvas.SetTop(CoordinatesDisplay, point.Y + 10);
                CoordinatesText.Text = $"X: {(int)point.X}, Y: {(int)point.Y}";
                CoordinatesDisplay.Visibility = Visibility.Visible;
            }
        }

        protected override void OnMouseLeftButtonUp(MouseButtonEventArgs e)
        {
            base.OnMouseLeftButtonUp(e);

            if (_isSelecting)
            {
                _isSelecting = false;
                Mouse.Capture(null);

                // Check if we have a valid selection
                if (_selectedRegion.Width > 5 && _selectedRegion.Height > 5)
                {
                    SelectionCancelled = false;

                    // Convert to screen coordinates
                    var screenPoint = PointToScreen(new System.Windows.Point(_selectedRegion.X, _selectedRegion.Y));
                    _selectedRegion = new Rect(screenPoint.X, screenPoint.Y, _selectedRegion.Width, _selectedRegion.Height);

                    // Close the window with success
                    DialogResult = true;
                    Close();
                }
                else
                {
                    // Reset if selection is too small
                    ResetSelection();
                }
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            base.OnKeyDown(e);

            if (e.Key == Key.Escape)
            {
                // Cancel selection
                SelectionCancelled = true;
                DialogResult = false;
                Close();
            }
            else if (e.Key == Key.Enter && _selectedRegion.Width > 0 && _selectedRegion.Height > 0)
            {
                // Confirm selection
                SelectionCancelled = false;

                // Convert to screen coordinates
                var screenPoint = PointToScreen(new System.Windows.Point(_selectedRegion.X, _selectedRegion.Y));
                _selectedRegion = new Rect(screenPoint.X, screenPoint.Y, _selectedRegion.Width, _selectedRegion.Height);

                DialogResult = true;
                Close();
            }
            else if (e.Key == Key.C && Keyboard.Modifiers == ModifierKeys.Control)
            {
                // Copy region coordinates to clipboard
                if (_selectedRegion.Width > 0 && _selectedRegion.Height > 0)
                {
                    var text = $"Region: X={_selectedRegion.X:F0}, Y={_selectedRegion.Y:F0}, Width={_selectedRegion.Width:F0}, Height={_selectedRegion.Height:F0}";
                    System.Windows.Clipboard.SetText(text);

                    // Show brief notification
                    ShowNotification("Coordinates copied to clipboard!");
                }
            }
        }

        private void UpdateCoordinatesDisplay(double x, double y, double width, double height)
        {
            // Position the display near the selection
            Canvas.SetLeft(CoordinatesDisplay, x + 5);
            Canvas.SetTop(CoordinatesDisplay, y - 30);

            // Update text
            CoordinatesText.Text = $"X: {(int)x}, Y: {(int)y} | W: {(int)width}, H: {(int)height}";
        }

        private void ResetSelection()
        {
            SelectionRectangle.Visibility = Visibility.Collapsed;
            CoordinatesDisplay.Visibility = Visibility.Collapsed;
            _selectedRegion = Rect.Empty;
        }

        private void ShowNotification(string message)
        {
            // Create a temporary notification
            var notification = new Border
            {
                Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(200, 0, 100, 0)),
                CornerRadius = new CornerRadius(5),
                Padding = new Thickness(15, 10, 15, 10),
                HorizontalAlignment = System.Windows.HorizontalAlignment.Center,
                VerticalAlignment = System.Windows.VerticalAlignment.Bottom,
                Margin = new Thickness(0, 0, 0, 100)
            };

            var text = new TextBlock
            {
                Text = message,
                Foreground = System.Windows.Media.Brushes.White,
                FontSize = 14,
                FontWeight = FontWeights.Bold
            };

            notification.Child = text;
            ((Grid)Content).Children.Add(notification);

            // Remove after 2 seconds
            var timer = new System.Windows.Threading.DispatcherTimer();
            timer.Interval = TimeSpan.FromSeconds(2);
            timer.Tick += (s, e) =>
            {
                ((Grid)Content).Children.Remove(notification);
                timer.Stop();
            };
            timer.Start();
        }

        /// <summary>
        /// Static method to show the region selector and return the selected region
        /// </summary>
        public static Rect? SelectRegion()
        {
            var selector = new RegionSelectorWindow();

            if (selector.ShowDialog() == true && !selector.SelectionCancelled)
            {
                return selector.SelectedRegion;
            }

            return null;
        }
    }
}