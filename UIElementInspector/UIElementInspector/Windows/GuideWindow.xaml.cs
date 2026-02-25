using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace UIElementInspector.Windows
{
    public partial class GuideWindow : Window
    {
        public GuideWindow()
        {
            InitializeComponent();
            HighlightNavButton(nav1);
        }

        private void NavButton_Click(object sender, RoutedEventArgs e)
        {
            if (sender is System.Windows.Controls.Button button && button.Tag is string sectionName)
            {
                // Find the target section
                var section = contentPanel.FindName(sectionName) as FrameworkElement;
                if (section != null)
                {
                    section.BringIntoView();
                }

                HighlightNavButton(button);
            }
        }

        private void HighlightNavButton(System.Windows.Controls.Button activeButton)
        {
            // Reset all nav buttons
            foreach (var child in ((StackPanel)navPanel).Children)
            {
                if (child is System.Windows.Controls.Button btn)
                {
                    btn.Background = System.Windows.Media.Brushes.Transparent;
                    btn.Foreground = new SolidColorBrush(System.Windows.Media.Color.FromRgb(236, 239, 241)); // #ECEFF1
                }
            }

            // Highlight active
            activeButton.Background = new SolidColorBrush(System.Windows.Media.Color.FromRgb(21, 101, 192)); // #1565C0
            activeButton.Foreground = System.Windows.Media.Brushes.White;
        }

        private void CloseGuide_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
