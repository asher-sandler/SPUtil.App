using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Input;
using System.Diagnostics;
using SPUtil.App.ViewModels;
using SPUtil.Infrastructure;

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

        // ── Folder navigation — no server calls, PagesViewModel just re-filters
        //    the already-loaded item list. Same UX pattern as Library101. ──────

        private void PagesRow_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGridRow row) return;
            if (row.Item is not SPFileData item) return;
            if (DataContext is not PagesViewModel vm) return;

            if (item.IsFolder)
                vm.NavigateToFolder(item.FullPath);
        }

        private void BtnPagesFolderUp_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not PagesViewModel vm) return;
            vm.NavigateUp();
        }

        private void BtnOpenPagesFolder_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string url && !string.IsNullOrEmpty(url))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not open folder: {ex.Message}",
                        "Open Folder", MessageBoxButton.OK, MessageBoxImage.Warning);
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
