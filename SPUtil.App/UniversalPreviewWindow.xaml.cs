using System.Windows;
using SPUtil.Infrastructure;

namespace SPUtil.App.Views
{
    public partial class UniversalPreviewWindow : Window
    {
        // Оставляем пустой конструктор для MVVM
        public UniversalPreviewWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Получаем кнопку из контекста данных
            if ((sender as FrameworkElement)?.DataContext is DialogButton btn)
            {
                // Выполняем действие
                btn.Action?.Invoke();
                
                // Если кнопка закрывающая — закрываем окно
                if (btn.IsCancel) 
                {
                    this.Close();
                }
            }
        }
    }
}