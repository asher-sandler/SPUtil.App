using System.Windows;

namespace SPUtil.Views
{
    // ИЗМЕНИТЕ Page на Window
    public partial class OperationInfoWindow : Window
    {
        public OperationInfoWindow()
        {
            InitializeComponent();
        }

        // Добавьте проверку Dispatcher, так как вы вызываете это из асинхронного StartCopyProcess
        public void UpdateMessage(string msg)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => UpdateMessage(msg));
                return;
            }
            TxtMessage.Text = msg;
        }
    }
}