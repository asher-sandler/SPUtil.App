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
			// Открывает системный браузер по умолчанию
			System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
			{
				FileName = e.Uri.AbsoluteUri,
				UseShellExecute = true
			});
			e.Handled = true;
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
    }
}