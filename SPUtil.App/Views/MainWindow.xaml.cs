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

		// ── Site tree search — jump to first match, then step through all matches ──

		private List<SPNode> _leftMatches = new();
		private int _leftMatchIndex = -1;

		private List<SPNode> _rightMatches = new();
		private int _rightMatchIndex = -1;

		private void TxtSearchLeft_TextChanged(object sender, TextChangedEventArgs e)
		{
			_leftMatches = FindMatches(TvLeft, TxtSearchLeft.Text);
			_leftMatchIndex = _leftMatches.Count > 0 ? 0 : -1;
			UpdateSearchUi(TvLeft, TxtSearchLeft, TxtSearchCountLeft, _leftMatches, _leftMatchIndex);
		}

		private void TxtSearchRight_TextChanged(object sender, TextChangedEventArgs e)
		{
			_rightMatches = FindMatches(TvRight, TxtSearchRight.Text);
			_rightMatchIndex = _rightMatches.Count > 0 ? 0 : -1;
			UpdateSearchUi(TvRight, TxtSearchRight, TxtSearchCountRight, _rightMatches, _rightMatchIndex);
		}

		private void BtnClearSearchLeft_Click(object sender, RoutedEventArgs e)
		{
			TxtSearchLeft.Text = string.Empty;
		}

		private void BtnClearSearchRight_Click(object sender, RoutedEventArgs e)
		{
			TxtSearchRight.Text = string.Empty;
		}

		private void BtnSearchNextLeft_Click(object sender, RoutedEventArgs e)
		{
			if (_leftMatches.Count == 0) return;
			_leftMatchIndex = (_leftMatchIndex + 1) % _leftMatches.Count;
			UpdateSearchUi(TvLeft, TxtSearchLeft, TxtSearchCountLeft, _leftMatches, _leftMatchIndex);
		}

		private void BtnSearchPrevLeft_Click(object sender, RoutedEventArgs e)
		{
			if (_leftMatches.Count == 0) return;
			_leftMatchIndex = (_leftMatchIndex - 1 + _leftMatches.Count) % _leftMatches.Count;
			UpdateSearchUi(TvLeft, TxtSearchLeft, TxtSearchCountLeft, _leftMatches, _leftMatchIndex);
		}

		private void BtnSearchNextRight_Click(object sender, RoutedEventArgs e)
		{
			if (_rightMatches.Count == 0) return;
			_rightMatchIndex = (_rightMatchIndex + 1) % _rightMatches.Count;
			UpdateSearchUi(TvRight, TxtSearchRight, TxtSearchCountRight, _rightMatches, _rightMatchIndex);
		}

		private void BtnSearchPrevRight_Click(object sender, RoutedEventArgs e)
		{
			if (_rightMatches.Count == 0) return;
			_rightMatchIndex = (_rightMatchIndex - 1 + _rightMatches.Count) % _rightMatches.Count;
			UpdateSearchUi(TvRight, TxtSearchRight, TxtSearchCountRight, _rightMatches, _rightMatchIndex);
		}

		/// <summary>
		/// Finds ALL nodes (in current display order — the list is never
		/// re-sorted or filtered) whose Title contains the typed text
		/// (case-insensitive). Tree is confirmed flat in this app (SPNode.Children
		/// exists in the model but is never populated anywhere) — no
		/// expand-ancestors logic needed.
		/// </summary>
		private List<SPNode> FindMatches(TreeView tree, string text)
		{
			if (string.IsNullOrEmpty(text)) return new List<SPNode>();

			return tree.Items.Cast<object>()
				.OfType<SPNode>()
				.Where(node => node.Title.Contains(text, StringComparison.OrdinalIgnoreCase))
				.ToList();
		}

		/// <summary>
		/// Updates the "N/M" counter, the search box's error border (no matches),
		/// and selects/scrolls to the current match (if any).
		/// </summary>
		private void UpdateSearchUi(TreeView tree, TextBox searchBox, TextBlock countLabel,
			List<SPNode> matches, int currentIndex)
		{
			if (string.IsNullOrEmpty(searchBox.Text))
			{
				searchBox.ClearValue(TextBox.BorderBrushProperty);
				countLabel.Text = string.Empty;
				return;
			}

			if (matches.Count == 0)
			{
				searchBox.BorderBrush = System.Windows.Media.Brushes.IndianRed;
				countLabel.Text = "0/0";
				return;
			}

			searchBox.ClearValue(TextBox.BorderBrushProperty);
			countLabel.Text = $"{currentIndex + 1}/{matches.Count}";

			SelectNode(tree, matches[currentIndex]);
		}

		private void SelectNode(TreeView tree, SPNode node)
		{
			if (tree.ItemContainerGenerator.ContainerFromItem(node) is TreeViewItem container)
			{
				container.IsSelected = true;
				container.BringIntoView();
			}
		}

		// ── Splitter between "Source details" and "Destination details" ──
		// GridSplitter itself has no percentage-based constraint in plain XAML —
		// MinWidth on a ColumnDefinition is always in pixels, so the 35% floor
		// has to be recalculated in code every time the window (and therefore
		// the combined width of these two columns) changes size.
		private void MainContentGrid_SizeChanged(object sender, SizeChangedEventArgs e)
		{
			double combined = ColSourceDetails.ActualWidth + ColDestinationDetails.ActualWidth;
			if (combined <= 0) return;

			double minWidth = combined * 0.35;
			ColSourceDetails.MinWidth = minWidth;
			ColDestinationDetails.MinWidth = minWidth;
		}
    }
}