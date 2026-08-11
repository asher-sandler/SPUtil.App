using System.Windows;

namespace SPUtil.Views
{
    /// <summary>
    /// Result dialog of a page copy. Replaces a plain MessageBox so the detailed
    /// report can be offered as a link rather than as a second, equally weighted
    /// button — the report matters when something went wrong, not every time.
    /// </summary>
    public partial class CopyCompleteWindow : Window
    {
        /// <summary>True when the user asked to see the detailed report.</summary>
        public bool ShowReportRequested { get; private set; }

        public CopyCompleteWindow(string message, string summary, string warning, bool hasProblems)
        {
            InitializeComponent();

            TxtMessage.Text = message;
            TxtSummary.Text = summary;
            TxtWarning.Text = warning;

            // The link is always available, but it only draws attention when there is
            // something to look at.
            if (hasProblems)
            {
                TxtReportLink.Text       = "⚠ View detailed report — some WebParts were not copied";
                TxtReportLink.Foreground = System.Windows.Media.Brushes.Firebrick;
                TxtReportLink.FontWeight = FontWeights.SemiBold;
            }
        }

        private void LnkReport_Click(object sender, RoutedEventArgs e)
        {
            ShowReportRequested = true;
            DialogResult = true;
        }

        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }
    }
}