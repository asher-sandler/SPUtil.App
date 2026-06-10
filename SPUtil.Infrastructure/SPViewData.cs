using Prism.Mvvm;

namespace SPUtil.Infrastructure
{
    public class SPViewData : BindableBase
    {
        public string Title { get; set; } = string.Empty;
        public string Id { get; set; } = string.Empty;
        public string ViewQuery { get; set; } = string.Empty;
        public string[]? ViewFields { get; set; }
        public bool DefaultView { get; set; }
        public string ServerRelativeUrl { get; set; } = string.Empty;

        // 2026-06-10: SchemaXml uses backing field so setting it also notifies
        // FormattedSchemaXml — otherwise the TextBox never updates when a view is selected.
        private string _schemaXml = string.Empty;
        public string SchemaXml
        {
            get => _schemaXml;
            set
            {
                if (SetProperty(ref _schemaXml, value))
                    RaisePropertyChanged(nameof(FormattedSchemaXml));
            }
        }

        // Добавляем это свойство для счетчиков и сумм
        public string Aggregations { get; set; } = string.Empty;

        // 2026-06-10: IsSelected — user checkbox for selecting views to copy.
        // All checks happen at copy time (CopySelectedViewsAsync), not on checkbox click.
        private bool _isSelected;
        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        // 2026-06-10: indented SchemaXml for display in the Views tab panel.
        // Notified automatically when SchemaXml changes.
        public string FormattedSchemaXml =>
            string.IsNullOrWhiteSpace(_schemaXml)
                ? string.Empty
                : SPUsingUtils.FormatXml(_schemaXml);
    }
}