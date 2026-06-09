using System;
using System.Globalization;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using SPUtil.App.ViewModels;
using SPUtil.Infrastructure;

namespace SPUtil.App.Views
{
    public partial class List100View : UserControl
    {
        public List100View()
        {
            InitializeComponent();
        }

        // 2026-06-09: routes checkbox click to ViewModel.OnViewCheckboxChanged.
        // WPF updates the binding before the Click event fires, so we pass the
        // new value (IsChecked) directly to let the VM validate and revert if needed.
        private void ViewCheckBox_Click(object sender, RoutedEventArgs e)
        {
            if (sender is CheckBox cb &&
                cb.DataContext is SPViewData view &&
                DataContext is List100ViewModel vm)
            {
                vm.OnViewCheckboxChanged(view, cb.IsChecked == true);
            }
        }
    }

    // 2026-06-09: simple bool inverter — used to disable checkboxes
    // for views that already exist on the target list (ExistsOnTarget = true).
    [ValueConversion(typeof(bool), typeof(bool))]
    public class InverseBoolConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && !b;

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
            => value is bool b && !b;
    }
}