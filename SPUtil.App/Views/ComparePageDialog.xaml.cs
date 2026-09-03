using System.Windows;

namespace SPUtil.Views
{
    public partial class ComparePageDialog : Window
    {
        /// <summary>Target page name without .aspx</summary>
        public string TargetPageName => TxtTargetPageName.Text.Trim()
            .Replace(".aspx", "", System.StringComparison.OrdinalIgnoreCase);


        /// <summary>
        /// Subfolder inside the Pages library, without the "Pages/" prefix.
        /// Empty means the library root. Nested paths are allowed ("Admin/Sub").
        /// </summary>
        public string TargetSubfolder => TxtTargetSubfolder.Text.Trim();

		public ComparePageDialog(
            string sourcePageName,
            string targetSiteUrl,
            string sourceInfo,
            string sourceSubfolder = "",
            string confirmButtonText = "Compare")
        {
            InitializeComponent();

            TxtTargetPageName.Text  = sourcePageName;
            TxtTargetUrl.Text       = targetSiteUrl;
            TxtInfo.Text            = sourceInfo;

            // Defaults to the source page's folder, so the previous behaviour is
            // preserved when the user does not touch the field.
            TxtTargetSubfolder.Text = sourceSubfolder;

            // Same dialog is reused for both "Compare" (ExecuteCompareWebPartAsync)
            // and "Copy WebPart Properties" (ExecuteCopyWebPartPropertiesAsync) —
            // the button previously always said "Compare" even when copying, which
            // was confusing. Defaults to "Compare" so the existing call site with
            // 4 arguments keeps working unchanged.
            BtnCompare.Content = confirmButtonText;

            TxtTargetPageName.Focus();
            TxtTargetPageName.SelectAll();
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
