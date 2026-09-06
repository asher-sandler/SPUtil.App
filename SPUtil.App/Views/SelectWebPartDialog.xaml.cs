using System.Collections.Generic;
using System.Windows;
using System.Windows.Input;
using SPUtil.Infrastructure;

namespace SPUtil.Views
{
    public partial class SelectWebPartDialog : Window
    {
        /// <summary>The candidate the user picked. Null unless DialogResult == true.</summary>
        public WebPartSnapshot? SelectedWebPart => GridCandidates.SelectedItem as WebPartSnapshot;

        public SelectWebPartDialog(List<WebPartSnapshot> candidates, string webPartTitle)
        {
            InitializeComponent();

            TxtHeader.Text =
                $"Found {candidates.Count} WebParts titled '{webPartTitle}' on the target page. " +
                "Select which one to use:";

            GridCandidates.ItemsSource = candidates;
            if (candidates.Count > 0)
                GridCandidates.SelectedIndex = 0;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            if (GridCandidates.SelectedItem == null)
            {
                MessageBox.Show("Please select a WebPart from the list.", "Validation",
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        private void GridCandidates_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            if (GridCandidates.SelectedItem != null)
                BtnOk_Click(sender, e);
        }
    }
}
