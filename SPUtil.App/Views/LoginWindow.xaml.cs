using SPUtil.Infrastructure;
using System;
using System.Windows;

namespace SPUtil.App.Views
{
    public partial class LoginWindow : Window
    {
        /// <summary>Account name without the domain prefix, as stored in the registry.</summary>
        public string UserName { get; private set; }

        public string Password { get; private set; }

        /// <summary>AD domain these credentials are being entered for.</summary>
        private readonly string _domain;

        /// <param name="domain">
        /// Domain resolved from the site URL. The two farms have no trust between them,
        /// so an account of the other domain is rejected rather than silently stored.
        /// </param>
        public LoginWindow(string domain = "")
        {
            InitializeComponent();

            _domain = (domain ?? string.Empty).Trim();

            if (string.IsNullOrEmpty(_domain))
            {
                PnlDomain.Visibility = Visibility.Collapsed;
            }
            else
            {
                TxtDomain.Text = $"Domain:  {_domain.ToUpperInvariant()}";
                Title = $"SharePoint Credentials — {_domain.ToUpperInvariant()}";
            }

            TxtUser.Focus();
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            HideError();

            if (string.IsNullOrWhiteSpace(TxtUser.Text) || string.IsNullOrWhiteSpace(TxtPass.Password))
                return;

            var (enteredDomain, user) = SPUsingUtils.SplitUserName(TxtUser.Text);

            // Only "user" and "DOMAIN\user" are accepted. UPN form, extra separators and
            // characters SharePoint does not allow in an account name are all rejected.
            bool valid =
                user.Length > 0 &&
                user.IndexOfAny(new[] { '@', '\\', '/', '[', ']', ':', ';', '|',
                                        '=', ',', '+', '*', '?', '<', '>', '"' }) < 0 &&
                (enteredDomain.Length == 0 ||
                 (!string.IsNullOrEmpty(_domain) &&
                  enteredDomain.Equals(_domain, StringComparison.OrdinalIgnoreCase)));

            if (!valid)
            {
                ShowError();
                return;
            }

            // Stored without the prefix: GetCredentials passes the domain to
            // NetworkCredential separately, and "DOMAIN\DOMAIN\user" would never
            // authenticate.
            UserName     = user;
            Password     = TxtPass.Password;
            DialogResult = true;
        }

        private void ShowError()
        {
            string d = string.IsNullOrEmpty(_domain) ? "this domain" : _domain.ToUpperInvariant();
            TxtError.Text = $"Enter the user name for domain {d}, " +
                            $"with or without the \"{d}\\\" prefix.";
            TxtError.Visibility = Visibility.Visible;

            TxtUser.Focus();
            TxtUser.SelectAll();
        }

        private void HideError() => TxtError.Visibility = Visibility.Collapsed;
    }
}