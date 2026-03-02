using System;
using System.Windows;

namespace UIElementInspector
{
    public partial class CaptureNameDialog : Window
    {
        public string CaptureName { get; private set; }

        public CaptureNameDialog(string defaultName = null)
        {
            InitializeComponent();
            NameTextBox.Text = defaultName ?? $"Capture_{DateTime.Now:yyyyMMdd_HHmmss}";

            // Force keyboard focus after window is fully activated and rendered
            this.Activated += (s, e) =>
            {
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(() =>
                {
                    NameTextBox.Focus();
                    System.Windows.Input.Keyboard.Focus(NameTextBox);
                    NameTextBox.SelectAll();
                }));
            };

            this.ContentRendered += (s, e) =>
            {
                Dispatcher.BeginInvoke(System.Windows.Threading.DispatcherPriority.Input, new Action(() =>
                {
                    this.Activate();
                    NameTextBox.Focus();
                    System.Windows.Input.Keyboard.Focus(NameTextBox);
                    NameTextBox.SelectAll();
                }));
            };
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            var name = NameTextBox.Text.Trim();
            if (string.IsNullOrEmpty(name))
            {
                System.Windows.MessageBox.Show("Lütfen bir isim girin.", "Uyarı", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Remove invalid file name characters
            foreach (var c in System.IO.Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            CaptureName = name;
            DialogResult = true;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }
    }
}
