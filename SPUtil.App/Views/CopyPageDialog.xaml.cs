using System;
using System.Linq;
using System.Windows;

namespace SPUtil.Views
{
    public partial class CopyPageDialog : Window
    {
        /// <summary>Characters SharePoint rejects in a folder name. The slash is a
        /// separator here and therefore allowed.</summary>
        private static readonly char[] InvalidFolderChars =
            { '~', '"', '#', '%', '&', '*', ':', '<', '>', '?', '\\', '{', '|', '}' };

        /// <summary>
        /// Remembers what the user typed while the checkbox is off, so unchecking and
        /// re-checking does not lose the path.
        /// </summary>
        private string _savedSubfolder = string.Empty;

        /// <summary>Target page name without .aspx</summary>
        public string TargetPageName => TxtTargetPageName.Text.Trim()
            .Replace(".aspx", "", System.StringComparison.OrdinalIgnoreCase);

        /// <summary>
        /// Folder inside the Pages library, without the "Pages/" prefix.
        /// Empty means the library root. Nested paths are allowed ("Admin/Sub").
        /// </summary>
        public string TargetSubfolder =>
            ChkKeepPath.IsChecked == true ? TxtSubfolder.Text.Trim().Trim('/') : string.Empty;

        /// <summary>
        /// Kept for callers that only need to know whether a folder was chosen.
        /// </summary>
        public bool KeepFolderPath => ChkKeepPath.IsChecked == true;

        /// <summary>The subfolder path passed in by the caller (read-only reference)</summary>
        public string SubfolderPath { get; private set; } = string.Empty;

        /// <param name="sourcePageName">Pre-fills the target name field</param>
        /// <param name="targetSiteUrl">Shown read-only</param>
        /// <param name="sourceInfo">Text shown in Source Page box</param>
        /// <param name="subfolderPath">
        /// Subfolder within Pages on the source site (e.g. "Dean"). Pre-fills the folder
        /// field; the user may change it to any other folder, existing or not.
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

            _savedSubfolder      = subfolderPath ?? string.Empty;
            TxtSubfolder.Text    = _savedSubfolder;

            // A page taken from the Pages root has nothing to keep, so the checkbox
            // starts off — but the field is still there if the user wants to place the
            // copy into a folder.
            ChkKeepPath.IsChecked = !string.IsNullOrEmpty(_savedSubfolder);

            ApplyPathState();
        }

        /// <summary>
        /// Caption of the confirmation button. The dialog is shared by the copy,
        /// rename and compare scenarios, so the caller adjusts the wording.
        /// Defaults to the value set in XAML ("Copy Page").
        /// </summary>
        public string ConfirmButtonText
        {
            get => BtnConfirm.Content?.ToString() ?? string.Empty;
            set => BtnConfirm.Content = value;
        }


        private void ChkKeepPath_Changed(object sender, RoutedEventArgs e)
        {
            // Fires during InitializeComponent when IsChecked="True" is applied from
            // XAML — at that point TxtSubfolder has not been created yet.
            if (TxtSubfolder == null) return;

            if (ChkKeepPath.IsChecked == true)
            {
                TxtSubfolder.Text = _savedSubfolder;
            }
            else
            {
                // Keep whatever was typed so re-checking restores it
                _savedSubfolder   = TxtSubfolder.Text.Trim();
                TxtSubfolder.Text = string.Empty;
            }

            ApplyPathState();
        }

        private void ApplyPathState()
        {
            bool on = ChkKeepPath.IsChecked == true;

            TxtSubfolder.IsEnabled  = on;
            TxtSubfolder.Background = on
                ? System.Windows.Media.Brushes.White
                : new System.Windows.Media.SolidColorBrush(
                      System.Windows.Media.Color.FromRgb(0xE9, 0xE9, 0xE9));

            UpdateHint();
        }

        private void UpdateHint()
        {
            string name   = string.IsNullOrWhiteSpace(TargetPageName) ? "page" : TargetPageName;
            string folder = TargetSubfolder;

            TxtPathHint.Text = string.IsNullOrEmpty(folder)
                ? $"Will create: Pages/{name}.aspx"
                : $"Will create: Pages/{folder}/{name}.aspx";
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrWhiteSpace(TargetPageName))
            {
                MessageBox.Show("Please enter a target page name.", "Validation",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string folder = TargetSubfolder;
            if (!string.IsNullOrEmpty(folder))
            {
                // The folder is created if it does not exist, so a typo silently produces
                // a new folder rather than an error. Only genuinely invalid input is
                // rejected here.
                string problem = ValidateFolder(folder);
                if (problem != null)
                {
                    MessageBox.Show(problem, "Invalid folder",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }
            }

            DialogResult = true;
            Close();
        }

        /// <summary>Returns an error message, or null when the path is acceptable.</summary>
        private static string ValidateFolder(string folder)
        {
            if (folder.IndexOfAny(InvalidFolderChars) >= 0)
                return "A folder name cannot contain any of:  ~ \" # % & * : < > ? \\ { | }";

            foreach (var segment in folder.Split('/'))
            {
                if (string.IsNullOrWhiteSpace(segment))
                    return "The path contains an empty folder name — check for double slashes.";

                if (segment.StartsWith("."))
                    return $"A folder name cannot start with a dot:  {segment}";

                if (segment.Length > 128)
                    return $"A folder name is too long (over 128 characters):  {segment}";

                if (segment.Equals("forms", StringComparison.OrdinalIgnoreCase))
                    return "'Forms' is reserved by SharePoint and cannot be used as a folder name.";
            }

            return null;
        }
    }
}