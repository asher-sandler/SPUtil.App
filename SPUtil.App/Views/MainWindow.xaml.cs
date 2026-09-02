using System;
using System.Linq;
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
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }
		private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
		{
			// e.Uri can be null if NavigateUri was empty string on first render
			string? url = e.Uri?.AbsoluteUri;
			if (string.IsNullOrEmpty(url)) return;

			System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
			{
				FileName = url,
				UseShellExecute = true
			});
			e.Handled = true;
		}

		private void Logo_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            var aboutWin = new AboutWindow
            {
                Owner = this // Указываем главное окно владельцем для корректного центрирования
            };
            aboutWin.ShowDialog();
        }		
		private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			// 1. Берем уже существующую ViewModel из DataContext окна
			if (this.DataContext is MainWindowViewModel vm)
			{
                // 2. Вызываем метод подтверждения выхода напрямую
                var result = MessageBox.Show(
                            "Are you sure you want to exit the application?",
                            "Confirm Exit",
                            MessageBoxButton.YesNo,
                            MessageBoxImage.Question);

                // Если пользователь передумал (нажал No) — отменяем закрытие
                if (result == MessageBoxResult.No)
                {
                    e.Cancel = true; // Да, я подтверждаю ОТМЕНУ закрытия
                }
                //else
                //{
                    
                    // e.Cancel = false; // Нет, я НЕ ОТМЕНЯЮ закрытие (пусть закрывается)
                    // это и так по умолчанию
                //}
                
			}
		}

		// ── Site tree search (minimal v1 — jump to first match, no next/prev yet) ──

		private void TxtSearchLeft_TextChanged(object sender, TextChangedEventArgs e)
		{
			SearchTree(TvLeft, TxtSearchLeft);
		}

		private void TxtSearchRight_TextChanged(object sender, TextChangedEventArgs e)
		{
			SearchTree(TvRight, TxtSearchRight);
		}

		private void BtnClearSearchLeft_Click(object sender, RoutedEventArgs e)
		{
			TxtSearchLeft.Text = string.Empty;
		}

		private void BtnClearSearchRight_Click(object sender, RoutedEventArgs e)
		{
			TxtSearchRight.Text = string.Empty;
		}

		/// <summary>
		/// Finds the first node (in current display order — the list is never
		/// re-sorted or filtered) whose Title contains the typed text
		/// (case-insensitive), and selects/scrolls to it. Tree is confirmed
		/// flat in this app (SPNode.Children exists in the model but is never
		/// populated anywhere) — no expand-ancestors logic needed.
		/// </summary>
		private void SearchTree(TreeView tree, TextBox searchBox)
		{
			string text = searchBox.Text;
			if (string.IsNullOrEmpty(text))
			{
				searchBox.ClearValue(TextBox.BorderBrushProperty);
				return;
			}

			var match = tree.Items.Cast<object>()
				.FirstOrDefault(item => item is SPNode node &&
					node.Title.Contains(text, StringComparison.OrdinalIgnoreCase));

			if (match == null)
			{
				searchBox.BorderBrush = System.Windows.Media.Brushes.IndianRed;
				return;
			}

			searchBox.ClearValue(TextBox.BorderBrushProperty);

			if (tree.ItemContainerGenerator.ContainerFromItem(match) is TreeViewItem container)
			{
				container.IsSelected = true;
				container.BringIntoView();
			}
		}
    }
}