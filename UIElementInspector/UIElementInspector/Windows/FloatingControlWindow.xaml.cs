using System;
using System.Windows;
using System.Windows.Input;

namespace UIElementInspector.Windows
{
    /// <summary>
    /// Floating control window that stays on top during inspection
    /// Allows users to control inspection without blocking UI elements
    /// </summary>
    public partial class FloatingControlWindow : Window
    {
        // Events to communicate with MainWindow
        public event EventHandler StopInspectionRequested;
        public event EventHandler ShowMainWindowRequested;

        public FloatingControlWindow()
        {
            InitializeComponent();

            // Register ESC key handler
            this.KeyDown += FloatingControlWindow_KeyDown;

            // Position window at top-right corner
            this.Loaded += FloatingControlWindow_Loaded;
        }

        private void FloatingControlWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // Position at center-top of screen for better visibility
            this.Left = (SystemParameters.PrimaryScreenWidth - this.Width) / 2;
            this.Top = 100;

            // Alternative: If you prefer top-right corner, uncomment below:
            // this.Left = SystemParameters.PrimaryScreenWidth - this.Width - 20;
            // this.Top = 20;
        }

        private void FloatingControlWindow_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key == System.Windows.Input.Key.F2)
            {
                StopInspection_Click(sender, new RoutedEventArgs());
            }
        }

        private void StopInspection_Click(object sender, RoutedEventArgs e)
        {
            // Notify MainWindow to stop inspection
            StopInspectionRequested?.Invoke(this, EventArgs.Empty);
        }

        private void ShowMain_Click(object sender, RoutedEventArgs e)
        {
            // Notify MainWindow to show itself
            ShowMainWindowRequested?.Invoke(this, EventArgs.Empty);
        }

        /// <summary>
        /// Update the mode display
        /// </summary>
        public void UpdateMode(string mode)
        {
            Dispatcher.Invoke(() =>
            {
                txtMode.Text = mode;
            });
        }

        /// <summary>
        /// Update the status display
        /// </summary>
        public void UpdateStatus(string status)
        {
            Dispatcher.Invoke(() =>
            {
                txtStatus.Text = status;
            });
        }

        /// <summary>
        /// Update the element count display
        /// </summary>
        public void UpdateElementCount(int count)
        {
            Dispatcher.Invoke(() =>
            {
                txtElementCount.Text = count.ToString();
            });
        }

        /// <summary>
        /// Prevent window from being closed, just hide it instead
        /// </summary>
        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            e.Cancel = true;
            this.Hide();
            base.OnClosing(e);
        }
    }
}
