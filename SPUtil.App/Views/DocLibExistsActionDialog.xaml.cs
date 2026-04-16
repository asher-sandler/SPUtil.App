using System.Windows;

namespace SPUtil.Views
{
    public partial class DocLibExistsActionDialog : Window
    {
        /// <summary>
        /// Selected action: "Append", "Overwrite", or "Mirror"
        /// </summary>
        public string SelectedAction { get; private set; } = "Cancel";

        public DocLibExistsActionDialog(string libraryName)
        {
            InitializeComponent();
            TxtLibName.Text = libraryName;
        }

        private void BtnContinue_Click(object sender, RoutedEventArgs e)
        {
            if (RadioAppend.IsChecked == true)
                SelectedAction = "Append";
            else if (RadioOverwrite.IsChecked == true)
                SelectedAction = "Overwrite";
            else if (RadioMirror.IsChecked == true)
                SelectedAction = "Mirror";

            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            SelectedAction = "Cancel";
            DialogResult = false;
            Close();
        }
    }
}
