using System.Windows;
using SPUtil.Infrastructure;

namespace SPUtil.App.Views
{
    public partial class UniversalPreviewWindow : Window
    {
        // Empty constructor for MVVM
        public UniversalPreviewWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            // Get the button from the data context
            if ((sender as FrameworkElement)?.DataContext is DialogButton btn)
            {
                // Execute the action
                btn.Action?.Invoke();
                
                // If this is a cancel/close button — close the window
                if (btn.IsCancel) 
                {
                    this.Close();
                }
            }
        }
    }
}