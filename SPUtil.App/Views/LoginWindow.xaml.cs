using System.Windows;
namespace SPUtil.App.Views
{
    public partial class LoginWindow : Window
    {
        public string UserName { get; private set; }
        public string Password { get; private set; }
        public LoginWindow() => InitializeComponent();
        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TxtUser.Text) || string.IsNullOrWhiteSpace(TxtPass.Password)) return;
            UserName = TxtUser.Text;
            Password = TxtPass.Password;
            DialogResult = true;
        }
    }
}