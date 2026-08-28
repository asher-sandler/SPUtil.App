using System.Windows;
using System.Windows.Controls;

namespace SPUtil.App.Controls
{
    // Reusable "please wait" overlay — semi-transparent layer with an
    // indeterminate progress bar. Place as the last child of a Grid that
    // also contains the DataGrid/panel it should cover, and bind IsBusy
    // to a ViewModel flag set around the slow await call.
    public partial class LoadingOverlay : UserControl
    {
        public static readonly DependencyProperty IsBusyProperty =
            DependencyProperty.Register(nameof(IsBusy), typeof(bool), typeof(LoadingOverlay),
                new PropertyMetadata(false));

        public static readonly DependencyProperty LoadingTextProperty =
            DependencyProperty.Register(nameof(LoadingText), typeof(string), typeof(LoadingOverlay),
                new PropertyMetadata("Loading data, please wait…"));

        public bool IsBusy
        {
            get => (bool)GetValue(IsBusyProperty);
            set => SetValue(IsBusyProperty, value);
        }

        public string LoadingText
        {
            get => (string)GetValue(LoadingTextProperty);
            set => SetValue(LoadingTextProperty, value);
        }

        public LoadingOverlay()
        {
            InitializeComponent();
        }
    }
}
