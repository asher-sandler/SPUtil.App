using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Serilog;
using SPUtil.Infrastructure;

namespace SPUtil.Views
{
    // Prototype only — no real CAML/service call. Dummy data used to
    // let the user test the Filter dialog UX (warning + cancellable
    // "query") before any real implementation is written.
    public partial class FilterDialog : Window
    {
        // True when the dummy "query" completed and produced a result —
        // caller (List100View) reads this to decide whether Reset Filter
        // should become enabled.
		public bool FilterApplied { get; private set; }

        // Exposed for the caller (List100ViewModel, once wired) to actually
        // run the query. Populated in BtnRequestData_Click after validation.
        public string? GeneratedCaml { get; private set; }

        public FilterDialog()
        {
            InitializeComponent();
            UpdateOperatorsForField();
        }

        private void CmbField_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            UpdateOperatorsForField();
        }


        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
		private void UpdateOperatorsForField()
		{
			if (CmbField == null || CmbOperator == null) return;

			string field = (CmbField.SelectedItem as ComboBoxItem)?.Content as string ?? "ID";
			CmbOperator.Items.Clear();

			if (field == "Title")
			{
				CmbOperator.Items.Add("Equals");
				CmbOperator.Items.Add("Contains");
			}
			else // ID, Modified
			{
				CmbOperator.Items.Add("Equals");
				CmbOperator.Items.Add("Greater than");
				CmbOperator.Items.Add("Less than");
				CmbOperator.Items.Add("Between");
			}

			CmbOperator.SelectedIndex = 0;

			// Swap the Value1 control type: DatePicker for Modified, plain TextBox otherwise.
			bool isDateField = field == "Modified";
			TxtValue1.Visibility = isDateField ? Visibility.Collapsed : Visibility.Visible;
			DtpValue1.Visibility = isDateField ? Visibility.Visible   : Visibility.Collapsed;

			// Field changed — re-evaluate the "To" pair for the (possibly reset) operator.
			UpdateValue2Visibility();
		}

		private void CmbOperator_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			UpdateValue2Visibility();
		}

		private void UpdateValue2Visibility()
		{
			if (CmbOperator == null || CmbField == null) return;

			bool isBetween  = (CmbOperator.SelectedItem as string) == "Between";
			bool isDateField = ((CmbField.SelectedItem as ComboBoxItem)?.Content as string) == "Modified";

			TxtToLabel.Visibility = isBetween ? Visibility.Visible : Visibility.Collapsed;
			TxtValue2.Visibility  = (isBetween && !isDateField) ? Visibility.Visible : Visibility.Collapsed;
			DtpValue2.Visibility  = (isBetween &&  isDateField) ? Visibility.Visible : Visibility.Collapsed;
		}
		private void BtnRequestData_Click(object sender, RoutedEventArgs e)
		{
			string field = (CmbField.SelectedItem as ComboBoxItem)?.Content as string ?? "ID";
			string op    = CmbOperator.SelectedItem as string ?? "Equals";
			bool isDateField = field == "Modified";
			bool isBetween   = op == "Between";

			string value1 = isDateField
				? DtpValue1.SelectedDate?.ToString("yyyy-MM-ddT00:00:00Z") ?? string.Empty
				: TxtValue1.Text.Trim();

			string? value2 = isBetween
				? (isDateField ? DtpValue2.SelectedDate?.ToString("yyyy-MM-ddT00:00:00Z") : TxtValue2.Text.Trim())
				: null;

			string whereClause;
			try
			{
				whereClause = CamlFilterBuilder.BuildWhereClause(field, op, value1, value2);
			}
			catch (ArgumentException ex)
			{
				MessageBox.Show(ex.Message, "Filter", MessageBoxButton.OK, MessageBoxImage.Warning);
				return;
			}

			Log.Debug("Filter — generated CAML: {Caml}", whereClause);

			var warn = MessageBox.Show(
				"This query may return a large number of items and take a long time to complete.\nContinue?",
				"Filter", MessageBoxButton.YesNo, MessageBoxImage.Warning);

			if (warn != MessageBoxResult.Yes) return;

			GeneratedCaml  = whereClause;
			FilterApplied  = true;
			DialogResult   = true;
			Close();
		}
    }
}