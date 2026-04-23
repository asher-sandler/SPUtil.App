using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Services;
using SPUtil.Infrastructure;
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using System.Linq;
using System.Diagnostics;
using System;
using Serilog;

namespace SPUtil.App.ViewModels
{
    /// <summary>
    /// Tabs available in List100View.
    /// Passed as CommandParameter to toolbar buttons so each handler
    /// knows which tab is active at the moment of the click.
    /// </summary>
    public enum ListTab
    {
        Items,
        Fields,
        Views
    }

    public class List100ViewModel : BindableBase
    {
        private static readonly ILogger _log = Log.ForContext<List100ViewModel>();

        private readonly ISharePointService _spService;

        // ── Data collections ─────────────────────────────────────────────────
        private ObservableCollection<SPListItemData> _items  = new();
        private ObservableCollection<SPFieldData>    _fields = new();
        private ObservableCollection<SPViewData>     _views  = new();

        public ObservableCollection<SPListItemData> Items  { get => _items;  set => SetProperty(ref _items,  value); }
        public ObservableCollection<SPFieldData>    Fields { get => _fields; set => SetProperty(ref _fields, value); }
        public ObservableCollection<SPViewData>     Views  { get => _views;  set => SetProperty(ref _views,  value); }

        // ── Scalar state ─────────────────────────────────────────────────────
        private string     _listTitle     = string.Empty;
        private string     _statusMessage = "Ready";
        private bool       _isSourceMode;
        private SPViewData _selectedView;

        public string     ListTitle     { get => _listTitle;     set => SetProperty(ref _listTitle,     value); }
        public string     StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public bool       IsSourceMode  { get => _isSourceMode;  set => SetProperty(ref _isSourceMode,  value); }
        public SPViewData SelectedView  { get => _selectedView;  set => SetProperty(ref _selectedView,  value); }

        // ── Active tab ───────────────────────────────────────────────────────
        // ActiveTab (enum) is the source of truth.
        // ActiveTabIndex (int) is what TabControl.SelectedIndex binds to —
        // changing the tab updates the enum automatically and vice-versa.
        private ListTab _activeTab = ListTab.Items;
        public ListTab ActiveTab
        {
            get => _activeTab;
            set
            {
                if (SetProperty(ref _activeTab, value))
                    RaisePropertyChanged(nameof(ActiveTabIndex));
            }
        }

        public int ActiveTabIndex
        {
            get => (int)_activeTab;
            set => ActiveTab = (ListTab)value;
        }

        // ── Stored context for Refresh ────────────────────────────────────────
        private string _lastSiteUrl  = string.Empty;
        private string _lastListPath = string.Empty;

        // ── Commands ─────────────────────────────────────────────────────────
        // DelegateCommand<object> — XAML passes CommandParameter="{Binding ActiveTab}"
        // so the handler receives the ListTab enum value at the moment of the click.

        public DelegateCommand<object> CreateOnTargetCommand { get; }
        public DelegateCommand<object> CopyWithDataCommand   { get; }
        public DelegateCommand<object> CopyViewsCommand      { get; }
        public DelegateCommand<object> CompareCommand        { get; }
        public DelegateCommand         RefreshCommand        { get; }

        public List100ViewModel(ISharePointService spService)
        {
            _spService = spService;

            CreateOnTargetCommand = new DelegateCommand<object>(param =>
            {
                var tab = ToTab(param);
                switch (tab)
                {
                    case ListTab.Items:
                        LogAndStatus("Create list structure on target site [tab: Items]");
                        break;
                    case ListTab.Fields:
                        LogAndStatus("Create list structure on target site [tab: Fields]");
                        break;
                    case ListTab.Views:
                        LogAndStatus("Create list structure on target site [tab: Views]");
                        break;
                }
            });

            CopyWithDataCommand = new DelegateCommand<object>(param =>
            {
                var tab = ToTab(param);
                switch (tab)
                {
                    case ListTab.Items:
                        LogAndStatus("Copy structure + data [tab: Items]");
                        break;
                    case ListTab.Fields:
                        LogAndStatus("Copy fields to target site [tab: Fields]");
                        break;
                    case ListTab.Views:
                        LogAndStatus("Copy views to target site [tab: Views]");
                        break;
                }
            });

            CopyViewsCommand = new DelegateCommand<object>(param =>
            {
                var tab = ToTab(param);
                LogAndStatus($"Copy views [active tab: {tab}]");
            });

            CompareCommand = new DelegateCommand<object>(param =>
            {
                var tab = ToTab(param);
                switch (tab)
                {
                    case ListTab.Items:
                        LogAndStatus("Compare items [tab: Items]");
                        break;
                    case ListTab.Fields:
                        LogAndStatus("Compare fields [tab: Fields]");
                        break;
                    case ListTab.Views:
                        LogAndStatus("Compare views [tab: Views]");
                        break;
                }
            });

            RefreshCommand = new DelegateCommand(async () =>
            {
                if (!string.IsNullOrEmpty(_lastSiteUrl) && !string.IsNullOrEmpty(_lastListPath))
                    await LoadDataAsync(_lastSiteUrl, _lastListPath);
            });
        }

        // ── Helpers ──────────────────────────────────────────────────────────

        /// <summary>
        /// Safely converts a CommandParameter (arrives as boxed enum or null) to ListTab.
        /// Falls back to the current ActiveTab if the cast fails.
        /// </summary>
        private ListTab ToTab(object param) =>
            param is ListTab t ? t : ActiveTab;

        private void LogAndStatus(string message)
        {
            StatusMessage = message;
            Debug.WriteLine($">>> [List100] {DateTime.Now:HH:mm:ss} — {message}");
        }

        // ── Data loading ─────────────────────────────────────────────────────
        public async Task LoadDataAsync(string siteUrl, string listPath)
        {
            _lastSiteUrl  = siteUrl;
            _lastListPath = listPath;

            LogAndStatus($"Loading list data: {listPath}...");
            Fields.Clear();
            Views.Clear();

            string cleanId = listPath.StartsWith("id:") ? listPath.Substring(3) : listPath;

            // ── Fields ──
            try
            {
                var fieldsData = await _spService.GetListFieldsAsync(siteUrl, cleanId);
                var result = fieldsData
                    .Where(f =>
                        (f.InternalName.StartsWith("_x") || !f.InternalName.StartsWith("_")) &&
                        f.TypeAsString != "Computed" &&
                        f.InternalName != "ContentTypeId" &&
                        f.InternalName != "Attachments")
                    .ToList();
                Fields = new ObservableCollection<SPFieldData>(result);
                LogAndStatus($"Fields loaded: {Fields.Count}");
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                LogAndStatus($"Field load error: {ex.Message}");
            }

            // ── Views ──
            try
            {
                var viewsData = await _spService.GetListViewsAsync(siteUrl, cleanId);
                Views = new ObservableCollection<SPViewData>(viewsData);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                Debug.WriteLine($"View load error: {ex.Message}");
            }

            // ── Items ──
            try
            {
                var allItems = await _spService.GetListItemsByIDAsync(siteUrl, cleanId);
                if (allItems.Count > 250)
                {
                    LogAndStatus($"Warning: list contains {allItems.Count} items. Showing first 250.");
                    Items = new ObservableCollection<SPListItemData>(allItems.Take(250));
                }
                else
                {
                    LogAndStatus($"Items: {allItems.Count}  |  Fields: {Fields.Count}  |  Views: {Views.Count}");
                    Items = new ObservableCollection<SPListItemData>(allItems);
                }
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                LogAndStatus($"Item load error: {ex.Message}");
            }
        }
    }
}
