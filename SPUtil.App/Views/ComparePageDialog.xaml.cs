using System.Windows;

namespace SPUtil.Views
{
    public partial class ComparePageDialog : Window
    {
        /// <summary>Target page name without .aspx</summary>
        public string TargetPageName => TxtTargetPageName.Text.Trim()
            .Replace(".aspx", "", System.StringComparison.OrdinalIgnoreCase);

        public ComparePageDialog(string sourcePageName, string targetSiteUrl, string sourceInfo)
        {
            InitializeComponent();
            TxtTargetPageName.Text = sourcePageName;
            TxtTargetUrl.Text      = targetSiteUrl;
            TxtInfo.Text           = sourceInfo;
        }

        private void BtnCompare_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TargetPageName))
            {
                MessageBox.Show("Please enter the target page name.", "Validation",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DialogResult = true;
            Close();
        }
    }
}
