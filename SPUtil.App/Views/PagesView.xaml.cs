using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;

namespace SPUtil.App.Views
{
    public partial class PagesView : UserControl
    {
        public PagesView()
        {
            InitializeComponent();
        }

        // Opens page-actions ContextMenu on left-click
        private void BtnHamburger_Click(object sender, RoutedEventArgs e)
        {
            OpenContextMenu(sender);
        }

        // Opens WP-actions ContextMenu on left-click
        private void BtnWpHamburger_Click(object sender, RoutedEventArgs e)
        {
            OpenContextMenu(sender);
        }

        private static void OpenContextMenu(object sender)
        {
            if (sender is Button btn && btn.ContextMenu != null)
            {
                btn.ContextMenu.PlacementTarget = btn;
                btn.ContextMenu.Placement = PlacementMode.Bottom;
                btn.ContextMenu.IsOpen = true;
            }
        }
    }
}
