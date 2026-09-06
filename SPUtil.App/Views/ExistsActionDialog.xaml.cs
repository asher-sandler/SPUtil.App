using System.Linq;
using System.Windows;
using System.Windows.Controls;

namespace SPUtil.Views
{
    public partial class ExistsActionDialog : Window
    {
        public string SelectedAction { get; private set; } = "Cancel";
        public string? NewName { get; private set; }

        public ExistsActionDialog(string listName)
        {
            InitializeComponent();
            TxtListName.Text = listName;
        }

        private void Action_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag != null)
            {
                SelectedAction = btn.Tag.ToString() ?? "Cancel";

                if (SelectedAction == "Rename")
                {
                    // Requires a reference to Microsoft.VisualBasic for Interaction.InputBox.
                    string input = Microsoft.VisualBasic.Interaction.InputBox(
                        "Enter new name for the target list:",
                        "Rename Target",
                        TxtListName.Text);

                    if (string.IsNullOrWhiteSpace(input) || input == TxtListName.Text)
                        return;

                    NewName = input;
                }

                this.DialogResult = SelectedAction != "Cancel";
                this.Close();
            }
        }
    }
}
