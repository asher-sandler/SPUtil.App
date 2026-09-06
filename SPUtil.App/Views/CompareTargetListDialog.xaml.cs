using System;
using System.Threading.Tasks;
using System.Windows;

namespace SPUtil.Views
{
    // Same "dumb dialog" convention as the rest of the app (e.g. CopyListDialog,
    // ExistsActionDialog) — no ISharePointService reference here. The existence
    // check needed to validate the typed name without closing the dialog is
    // injected as a delegate by the caller (MainWindowViewModel.CompareList),
    // which already owns the ISharePointService instance and the target site URL.
    public partial class CompareTargetListDialog : Window
    {
        private readonly Func<string, Task<bool>> _listExistsChecker;

        public string TargetListTitle => TxtTargetListName.Text.Trim();

        public CompareTargetListDialog(string defaultListTitle, string targetUrlTitle,
            Func<string, Task<bool>> listExistsChecker)
        {
            InitializeComponent();
            TxtTargetListName.Text = defaultListTitle;
            TxtTargetUrlName.Text = targetUrlTitle;
            _listExistsChecker = listExistsChecker;
        }

        private async void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            TxtError.Visibility = Visibility.Collapsed;

            if (string.IsNullOrWhiteSpace(TargetListTitle))
            {
                TxtError.Text = "Please enter a list name.";
                TxtError.Visibility = Visibility.Visible;
                return;
            }

            BtnOk.IsEnabled = false;
            try
            {
                bool exists = await _listExistsChecker(TargetListTitle);
                if (!exists)
                {
                    TxtError.Text = "List not found on the destination site. Check the name and try again.";
                    TxtError.Visibility = Visibility.Visible;
                    return;
                }
            }
            finally
            {
                BtnOk.IsEnabled = true;
            }

            DialogResult = true;
            Close();
        }
    }
}
