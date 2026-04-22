using System.Windows;

namespace SPUtil.Views
{
    public partial class CopyPageDialog : Window
    {
        /// <summary>Target page name without .aspx</summary>
        public string TargetPageName => TxtTargetPageName.Text.Trim()
            .Replace(".aspx", "", System.StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// True if user wants to preserve the subfolder path on target.
        /// Only relevant when SubfolderPath is non-empty.
        /// </summary>
        public bool KeepFolderPath => ChkKeepPath.IsChecked == true;

        /// <summary>The subfolder path passed in by the caller (read-only reference)</summary>
        public string SubfolderPath { get; private set; } = string.Empty;

        /// <param name="sourcePageName">Pre-fills the target name field</param>
        /// <param name="targetSiteUrl">Shown read-only</param>
        /// <param name="sourceInfo">Text shown in Source Page box</param>
        /// <param name="subfolderPath">
        /// Subfolder within Pages on the source site (e.g. "Dean").
        /// Empty = page is in Pages root, path checkbox is hidden.
        /// </param>
        public CopyPageDialog(
            string sourcePageName,
            string targetSiteUrl,
            string sourceInfo,
            string subfolderPath = "")
        {
            InitializeComponent();
            TxtTargetPageName.Text = sourcePageName;
            TxtTargetUrl.Text      = targetSiteUrl;
            TxtInfo.Text           = sourceInfo;
            SubfolderPath          = subfolderPath;

            if (!string.IsNullOrEmpty(subfolderPath))
            {
                // Show path option with hint
                PnlPathOption.Visibility = Visibility.Visible;
                TxtPathHint.Text =
                    $"Checked  → create Pages/{subfolderPath}/{sourcePageName}.aspx\n" +
                    $"Unchecked → create Pages/{sourcePageName}.aspx (root)";
                Height = 460;   // expand window to fit the panel
            }
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
