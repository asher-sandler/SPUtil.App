using System.Windows;

namespace SPUtil.Views
{
    public partial class CopyListDialog : Window
    {
        public string TargetListTitle => TxtTargetListName.Text;

        public CopyListDialog(string targetListTitle,string targetURLTitle, string detailedInfo)
        {
            InitializeComponent();
            TxtTargetListName.Text = targetListTitle;
            TxtTargetUrlName.Text = targetURLTitle;
            TxtInfo.Text = detailedInfo;
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            //MessageBox.Show("Triggered from: CopyListDialog.BtnCopy_Click");
            
            if (string.IsNullOrWhiteSpace(TargetListTitle))
            {
                MessageBox.Show("Please enter a target name.");
                return;
            }
            this.DialogResult = true;
            this.Close();
        }
    }
}