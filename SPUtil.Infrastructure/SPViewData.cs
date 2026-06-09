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
        public string SchemaXml { get; set; } = string.Empty;

        // Добавляем это свойство для счетчиков и сумм
        public string Aggregations { get; set; } = string.Empty;

        // 2026-06-09: added for Views-tab copy workflow.
        // IsSelected   — user checkbox; only enabled when !ExistsOnTarget.
        // ExistsOnTarget — set by LoadViewsStatusAsync once per tab open;
        //                  true = view already exists on target, checkbox disabled.
        private bool _isSelected;
        private bool _existsOnTarget;

        public bool IsSelected
        {
            get => _isSelected;
            set => SetProperty(ref _isSelected, value);
        }

        public bool ExistsOnTarget
        {
            get => _existsOnTarget;
            set => SetProperty(ref _existsOnTarget, value);
        }
    }
}