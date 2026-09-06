using System.Windows;
using SPUtil.Infrastructure;

namespace SPUtil.Views
{
    // Stub for now — shows which WebPart was targeted, no editing yet.
    // Editing logic (skip-list filtering, UpdateWebPartAsync, checkout/checkin)
    // will be added once the property categorization (A/B/C) is settled.
    public partial class CustomPropertiesEditorDialog : Window
    {
        public CustomPropertiesEditorDialog(SPWebPartData webPart)
        {
            InitializeComponent();
            TxtWpTitle.Text = $"WebPart: {webPart.Title}";
            TxtWpStorageKey.Text = $"StorageKey: {webPart.StorageKey}";
        }

        private void BtnClose_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}