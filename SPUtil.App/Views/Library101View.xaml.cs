﻿using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using SPUtil.App.ViewModels;
using SPUtil.Infrastructure;

namespace SPUtil.App.Views
{
    /// <summary>
    /// Interaction logic for Library101View.xaml
    /// </summary>
    public partial class Library101View : UserControl
    {
        public Library101View()
        {
            InitializeComponent();
        }

        private async void Row_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (sender is not DataGridRow row) return;
            if (row.Item is not SPFileData item) return;
            if (DataContext is not Library101ViewModel vm) return;

            // Only folders navigate — double-clicking a file row does nothing
            // for now (no "open file" action defined yet).
            if (item.IsFolder)
                await vm.NavigateToFolderAsync(item.FullPath);
        }

        private async void BtnFolderUp_Click(object sender, RoutedEventArgs e)
        {
            if (DataContext is not Library101ViewModel vm) return;
            await vm.NavigateUpAsync();
        }

        private void BtnOpenFolder_Click(object sender, RoutedEventArgs e)
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
    }
}