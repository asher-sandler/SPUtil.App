
using System.Windows;
using System.Windows.Media;

namespace SPUtil.Views
{
    public partial class CopyCompleteWindow : Window
    {
        public bool ShowReportRequested { get; private set; }

        public CopyCompleteWindow(string message, string summary, string warning, bool hasProblems)
        {
            InitializeComponent();

            TxtMessage.Text = message;
            TxtSummary.Text = summary;
            TxtWarning.Text = warning;

            if (hasProblems)
            {
                TxtReportLink.Text       = "⚠ View detailed report — some WebParts were not copied";
                TxtReportLink.Foreground = Brushes.Firebrick;
                TxtReportLink.FontWeight = FontWeights.SemiBold;

                // Make the summary box itself impossible to miss at a glance —
                // the badge plus the red tint, not just the small link below.
                IconErrorBadge.Visibility = Visibility.Visible;
                BorderSummary.Background  = new SolidColorBrush(Color.FromRgb(0xFD, 0xEC, 0xEA));
                BorderSummary.BorderBrush = new SolidColorBrush(Color.FromRgb(0xB0, 0x00, 0x20));
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