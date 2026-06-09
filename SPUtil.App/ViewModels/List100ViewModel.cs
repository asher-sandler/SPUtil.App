using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Services;
using SPUtil.Infrastructure;
using SPUtil.App.Views;
using SPUtil.Views;
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
                {
                    RaisePropertyChanged(nameof(ActiveTabIndex));
                    // 2026-06-09: re-evaluate button enabled state on every tab switch
                    CopyWithDataCommand.RaiseCanExecuteChanged();
                    CopyViewsCommand.RaiseCanExecuteChanged();
                }
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

        // ── Target site URL (set by MainWindowViewModel, same pattern as PagesViewModel) ──
        private string _targetSiteUrl = string.Empty;

        /// <summary>
        /// Called by MainWindowViewModel right after resolving this VM.
        /// Always the right-pane site URL — toolbar is hidden when IsSourceMode=false,
        /// so copy commands can only ever fire from the left pane.
        /// </summary>
        public void SetTargetSiteUrl(string targetSiteUrl) =>
            _targetSiteUrl = targetSiteUrl ?? string.Empty;

        // ── Commands ─────────────────────────────────────────────────────────
        // DelegateCommand<object> — XAML passes CommandParameter="{Binding ActiveTab}"
        // so the handler receives the ListTab enum value at the moment of the click.

        // 2026-06-09: CreateOnTargetCommand removed — creation is handled via main toolbar
        public DelegateCommand<object> CopyWithDataCommand   { get; }
        public DelegateCommand<object> CopyViewsCommand      { get; }
        // 2026-06-09: CompareCommand removed — duplicate of MainWindow CompareListsCommand
        public DelegateCommand         RefreshCommand        { get; }

        public List100ViewModel(ISharePointService spService)
        {
            _spService = spService;

            // 2026-06-09: CreateOnTargetCommand removed — creation is handled via main toolbar

            // 2026-06-09: CanExecute added — button is greyed out when tab context makes it irrelevant
            // Copy is only meaningful on Items tab (selected rows to copy)
            CopyWithDataCommand = new DelegateCommand<object>(
                executeMethod: async param =>
                {
                    var tab = ToTab(param);
                    switch (tab)
                    {
                        case ListTab.Items:
                            await CopySelectedItemsAsync();
                            break;
                        case ListTab.Fields:
                            LogAndStatus("Copy fields to target site [tab: Fields]");
                            break;
                        case ListTab.Views:
                            LogAndStatus("Copy views to target site [tab: Views]");
                            break;
                    }
                },
                canExecuteMethod: param => ToTab(param) == ListTab.Items);

            // CopyViews is only meaningful on Views tab
            CopyViewsCommand = new DelegateCommand<object>(
                executeMethod: param =>
                {
                    var tab = ToTab(param);
                    LogAndStatus($"Copy views [active tab: {tab}]");
                },
                canExecuteMethod: param => ToTab(param) == ListTab.Views);

            // 2026-06-09: CompareCommand removed — duplicate of MainWindow CompareListsCommand.
            // Fields tab: identical schema comparison already available via GetListSchemaAsync.
            // Items/Views tab: no implementation existed — only LogAndStatus stubs.
            // Use the Compare button in the main toolbar instead.

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

        /// <summary>
        /// Collects checked items (IsSelected == true), verifies the target list
        /// exists, warns the user that data will be appended, then hands off to
        /// the copy service.
        /// </summary>
        private async Task CopySelectedItemsAsync()
        {
            // ── 1. Collect selected items ────────────────────────────────────
            var selectedItems = Items.Where(i => i.IsSelected).ToList();
            if (selectedItems.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "No items selected.\nCheck at least one row in the Items tab.",
                    "Nothing Selected",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
                return;
            }

            // ── 2. Validate target site URL is known ─────────────────────────
            if (string.IsNullOrEmpty(_targetSiteUrl))
            {
                System.Windows.MessageBox.Show(
                    "Target site URL is not set. Please connect to the target site first.",
                    "No Target Site",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            // ── 3. Check target list existence ───────────────────────────────
            bool targetExists;
            try
            {
                targetExists = await _spService.ListExistsAsync(_targetSiteUrl, _listTitle);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR checking target list existence: {ExType} — {Message}",
                    ex.GetType().Name, ex.Message);
                System.Windows.MessageBox.Show(
                    $"Could not check target list:\n{ex.Message}",
                    "Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                return;
            }

            if (!targetExists)
            {
                System.Windows.MessageBox.Show(
                    $"List \"{_listTitle}\" does not exist on the target site:\n{_targetSiteUrl}\n\n" +
                    "You can create it first using the left panel menu\n" +
                    "(select the list → Copy structure to target).",
                    "List Not Found",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            // ── 4. Confirm: data will be appended ────────────────────────────
            var idList = string.Join(", ", selectedItems.Select(i => i.Id));
            var confirm = System.Windows.MessageBox.Show(
                $"{selectedItems.Count} item(s) will be APPENDED to\n" +
                $"list \"{_listTitle}\" on:\n{_targetSiteUrl}\n\n" +
                //$"IDs: {idList}\n\n" +
                "Existing items on the target will NOT be modified.\nContinue?",
                "Confirm Append",
                System.Windows.MessageBoxButton.OKCancel,
                System.Windows.MessageBoxImage.Question);

            if (confirm != System.Windows.MessageBoxResult.OK)
                return;

            // ── DEBUG ────────────────────────────────────────────────────────
            Debug.WriteLine($">>> [List100] CopySelectedItemsAsync — IDs: [{idList}]");
            _log.Information("CopySelectedItemsAsync — count: {Count}, IDs: [{Ids}]",
                selectedItems.Count, idList);
            LogAndStatus($"Copying {selectedItems.Count} item(s)...");

            // 2026-06-09: wired up real service call with ProgressWindow + CancellationToken
            using var cts = new System.Threading.CancellationTokenSource();

            var progressWin = new ProgressWindow(cts)
            {
                Owner = System.Windows.Application.Current.MainWindow
            };

            var progressIndicator = new Progress<CopyProgressArgs>(e =>
                progressWin.UpdateStatus(e.Processed, e.Total, e.Message));

            try
            {
                progressWin.Show();

                await _spService.CopyListItemsAsync(
                    sourceUrl     : _lastSiteUrl,
                    targetUrl     : _targetSiteUrl,
                    sourceTitle   : _listTitle,
                    targetListName: _listTitle,
                    action        : "Append",
                    progress      : progressIndicator,
                    ct            : cts.Token,
                    itemIds       : selectedItems.Select(i => i.Id));

                progressWin.Close();
                LogAndStatus($"Done. {selectedItems.Count} item(s) copied to {_targetSiteUrl}");
                _log.Information("CopySelectedItemsAsync complete — {Count} items", selectedItems.Count);
                System.Windows.MessageBox.Show(
                    $"{selectedItems.Count} item(s) successfully copied to:\n{_targetSiteUrl}\n\n" +
                    "Refresh the right panel to see the changes.",
                    "Copy Complete",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
            }
            catch (OperationCanceledException)
            {
                progressWin.Close();
                LogAndStatus("Copy cancelled.");
                _log.Warning("CopySelectedItemsAsync cancelled by user");
            }
            catch (Exception ex)
            {
                progressWin.Close();
                _log.Error(ex, "ERROR in CopySelectedItemsAsync: {ExType} — {Message}",
                    ex.GetType().Name, ex.Message);
                System.Windows.MessageBox.Show(
                    $"Copy error:\n{ex.Message}",
                    "Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                LogAndStatus($"Copy error: {ex.Message}");
            }
        }

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
