using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Diagnostics;

namespace SPUtil.App.Views
{
    public partial class PagesView : UserControl
    {
        public PagesView()
        {
            InitializeComponent();
        }
        private void BtnOpenPage_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string url && !string.IsNullOrEmpty(url))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not open page: {ex.Message}",
                        "Open Page", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        // Opens page-actions ContextMenu on left-click
        private void BtnHamburger_Click(object sender, RoutedEventArgs e)
        {
            OpenContextMenu(sender);
        }

        // Opens WP-actions ContextMenu on left-click
        private void BtnWpHamburger_Click(object sender, RoutedEventArgs e)
        {
            OpenContextMenu(sender);
        }

        private static void OpenContextMenu(object sender)
        {
            if (sender is Button btn && btn.ContextMenu != null)
            {
                btn.ContextMenu.PlacementTarget = btn;
                btn.ContextMenu.Placement = PlacementMode.Bottom;
                btn.ContextMenu.IsOpen = true;
            }
        }
    }
}
