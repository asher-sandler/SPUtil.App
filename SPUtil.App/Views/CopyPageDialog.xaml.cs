using System.Windows;

namespace SPUtil.Views
{
    public partial class CopyPageDialog : Window
    {
        /// <summary>Target page name without .aspx</summary>
        public string TargetPageName => TxtTargetPageName.Text.Trim()
            .Replace(".aspx", "", System.StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// "Replace"  — delete existing page and create from snapshot.
        /// "Rename"   — rename existing to Name_old, then create.
        /// </summary>
        public string ExistsAction => RadioRename.IsChecked == true ? "Rename" : "Replace";

        public CopyPageDialog(string sourcePageName, string targetSiteUrl, string sourceInfo)
        {
            InitializeComponent();
            TxtTargetPageName.Text = sourcePageName;
            TxtTargetUrl.Text      = targetSiteUrl;
            TxtInfo.Text           = sourceInfo;
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TargetPageName))
            {
                MessageBox.Show("Please enter a target page name.", "Validation",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DialogResult = true;
            Close();
        }
    }
}
