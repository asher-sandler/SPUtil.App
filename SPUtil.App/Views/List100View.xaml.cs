using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;
using System.Windows.Controls;
using SPUtil.App.ViewModels;
using SPUtil.Views;

namespace SPUtil.App.Views
{
    public partial class List100View : UserControl
    {
        public List100View()
        {
            InitializeComponent();
        }
		private void BtnOpenItem_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string url && !string.IsNullOrEmpty(url))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not open item: {ex.Message}",
                        "Open Item", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }

        private void BtnOpenList_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string url && !string.IsNullOrEmpty(url))
            {
                try
                {
                    Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Could not open list: {ex.Message}",
                        "Open List", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
            }
        }
        private async void BtnFilter_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new FilterDialog { Owner = Window.GetWindow(this) };
            bool? result = dialog.ShowDialog();

            if (result != true || !dialog.FilterApplied || string.IsNullOrEmpty(dialog.GeneratedCaml))
                return;

            if (DataContext is not List100ViewModel vm) return;

            // See caveat in FilterDialog/service comments: GetListItemsByIDAsync(whereClause)
            // has no CancellationToken — Cancel here only stops waiting on the UI side,
            // it does not abort the CSOM call already running on the background thread.
            using var cts = new CancellationTokenSource();
            var progress = new ProgressWindow(cts) { Owner = Window.GetWindow(this) };
            progress.Show();
            progress.UpdateStatus(0, 0, "Loading filtered items...");

            try
            {
                await vm.ApplyFilterAsync(dialog.GeneratedCaml);
                BtnResetFilter.IsEnabled = true;
            }
            finally
            {
                progress.Close();
            }
        }

        private async void BtnResetFilter_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not List100ViewModel vm) return;

            await vm.ResetFilterAsync();
            BtnResetFilter.IsEnabled = false;
        }
    }
}