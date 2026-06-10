using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Services;
using SPUtil.Infrastructure;
using SPUtil.App.Views;
using SPUtil.Views;
using System.Collections.Generic;
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
        private string      _listTitle     = string.Empty;
        private string      _statusMessage = "Ready";
        private bool        _isSourceMode;
        private SPViewData? _selectedView;

        public string      ListTitle     { get => _listTitle;     set => SetProperty(ref _listTitle,     value); }
        public string      StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public bool        IsSourceMode  { get => _isSourceMode;  set => SetProperty(ref _isSourceMode,  value); }
        public SPViewData? SelectedView  { get => _selectedView;  set => SetProperty(ref _selectedView,  value); }

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

                    // 2026-06-09: when user switches to Views tab, load target status
                    // in the background so checkboxes reflect availability immediately.
                    // Fire-and-forget — errors handled inside LoadViewsStatusAsync.
                    if (value == ListTab.Views)
                        _ = LoadViewsStatusAsync();
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

        // 2026-06-09: cached target-side state for Views tab.
        // Populated once by LoadViewsStatusAsync when the user switches to Views tab.
        // Used by the checkbox click handler to give instant feedback without network calls.
        private bool _targetListExists     = false;
        private bool _targetSchemaMatch    = false;
        private HashSet<string> _targetViewTitles = new(StringComparer.OrdinalIgnoreCase);

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

            // 2026-06-09: CopyViews wired to CopySelectedViewsAsync.
            // CanExecute keeps the button grey on other tabs.
            // CopyViews is only meaningful on Views tab
            CopyViewsCommand = new DelegateCommand<object>(
                executeMethod: async param =>
                {
                    var tab = ToTab(param);
                    if (tab == ListTab.Views)
                        await CopySelectedViewsAsync();
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

        // 2026-06-09: ── Views-tab status loading ────────────────────────────
        // Called fire-and-forget when user switches to Views tab.
        // Runs two network requests (ListExists + GetListViews) to populate
        // the three cached flags used by OnViewCheckboxChanged.
        private async Task LoadViewsStatusAsync()
        {
            // Reset all flags and checkboxes before every load
            _targetListExists  = false;
            _targetSchemaMatch = false;
            _targetViewTitles.Clear();

            foreach (var v in Views)
            {
                v.ExistsOnTarget = false;
                v.IsSelected     = false;
            }

            if (string.IsNullOrEmpty(_targetSiteUrl))
            {
                LogAndStatus("Views: connect to target site to see copy availability.");
                return;
            }

            LogAndStatus("Views: checking target...");

            // ── Step 1: does the list exist on target? ───────────────────────
            try
            {
                _targetListExists = await _spService.ListExistsAsync(_targetSiteUrl, _listTitle);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "LoadViewsStatusAsync — ListExistsAsync failed: {Message}", ex.Message);
                LogAndStatus("Views: could not reach target site.");
                return;
            }

            if (!_targetListExists)
            {
                LogAndStatus($"Views: list \"{_listTitle}\" not found on target — copy unavailable.");
                return;
            }

            // ── Step 2: compare field schemas source vs target ───────────────
            // Views contain CAML that references InternalNames. If schemas differ
            // the copied view may be broken, so we block copy when they diverge.
            // 2026-06-09 fix: apply the same filter to targetFields that LoadDataAsync
            // uses when building Fields — without it, raw system fields on target
            // (e.g. _UIVersionString, _ModerationStatus) inflate targetNames and
            // make IsSubsetOf unreliable. Both sides must be filtered identically.
            try
            {
                var targetFields = await _spService.GetListFieldsAsync(_targetSiteUrl, _listTitle);

                var sourceNames = Fields.Select(f => f.InternalName)
                                        .ToHashSet(StringComparer.OrdinalIgnoreCase);

                var targetNames = targetFields
                    .Where(f =>
                        (f.InternalName.StartsWith("_x") || !f.InternalName.StartsWith("_")) &&
                        f.TypeAsString != "Computed" &&
                        f.InternalName != "ContentTypeId" &&
                        f.InternalName != "Attachments")
                    .Select(f => f.InternalName)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                // Schema matches when every source field exists on target
                _targetSchemaMatch = sourceNames.IsSubsetOf(targetNames);

                if (!_targetSchemaMatch)
                {
                    var missingFields = sourceNames.Except(targetNames).ToList();
                    _log.Warning("LoadViewsStatusAsync — schema mismatch, missing on target: {Fields}",
                        string.Join(", ", missingFields));
                }
            }
            catch (Exception ex)
            {
                _log.Error(ex, "LoadViewsStatusAsync — GetListFieldsAsync failed: {Message}", ex.Message);
                LogAndStatus("Views: could not compare field schemas.");
                return;
            }

            if (!_targetSchemaMatch)
            {
                LogAndStatus("Views: field schema mismatch — copy unavailable. Run Compare from the left panel menu.");
                return;
            }

            // ── Step 3: load target views, build title set ───────────────────
            try
            {
                var targetViews = await _spService.GetListViewsAsync(_targetSiteUrl, _listTitle);
                _targetViewTitles = targetViews
                    .Select(v => v.Title)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "LoadViewsStatusAsync — GetListViewsAsync failed: {Message}", ex.Message);
                LogAndStatus("Views: could not load target views.");
                return;
            }

            // ── Step 4: mark each source view ────────────────────────────────
            int missing = 0;
            foreach (var v in Views)
            {
                v.ExistsOnTarget = _targetViewTitles.Contains(v.Title);
                v.IsSelected     = !v.ExistsOnTarget; // pre-check missing ones
                if (!v.ExistsOnTarget) missing++;
            }

            LogAndStatus(missing > 0
                ? $"Views: {missing} missing on target — ready to copy."
                : "Views: all views already exist on target.");

            _log.Information(
                "LoadViewsStatusAsync — source: {Total}, missing on target: {Missing}",
                Views.Count, missing);
        }

        // 2026-06-09: ── Checkbox click guard ────────────────────────────────
        // Called from XAML code-behind when user clicks a view checkbox.
        // Validates the cached flags and blocks the check with a message if needed.
        // No network calls here — all state was loaded by LoadViewsStatusAsync.
        public void OnViewCheckboxChanged(SPViewData view, bool newValue)
        {
            if (!newValue) return; // unchecking is always allowed

            if (!_targetListExists)
            {
                view.IsSelected = false;
                System.Windows.MessageBox.Show(
                    $"List \"{_listTitle}\" does not exist on the target site.\n\n" +
                    "Create it first using the left panel menu.",
                    "Cannot Copy View",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (!_targetSchemaMatch)
            {
                view.IsSelected = false;
                System.Windows.MessageBox.Show(
                    "The field schema on the target list differs from the source.\n\n" +
                    "View CAML queries reference field names that may not exist on target.\n" +
                    "Copy the list structure first.",
                    "Cannot Copy View",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            if (view.ExistsOnTarget)
            {
                view.IsSelected = false;
                System.Windows.MessageBox.Show(
                    $"View \"{view.Title}\" already exists on the target list.\n\n" +
                    "Only missing views can be copied.",
                    "Cannot Copy View",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            // All checks passed — allow the check
            view.IsSelected = true;
        }

        // 2026-06-09: ── Copy selected (missing) views to target ─────────────
        private async Task CopySelectedViewsAsync()
        {
            // 2026-06-09: cached flags (_targetListExists, _targetSchemaMatch) are for UX only.
            // Before the real operation we re-verify everything live — the target could have
            // changed since the tab was opened (list deleted, fields modified, etc.).

            LogAndStatus("Verifying target before copy...");

            // ── Live check 1: target site URL ────────────────────────────────
            if (string.IsNullOrEmpty(_targetSiteUrl))
            {
                System.Windows.MessageBox.Show(
                    "Target site URL is not set. Please connect to the target site first.",
                    "No Target Site",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            // ── Live check 2: list exists on target ──────────────────────────
            bool listExists;
            try
            {
                listExists = await _spService.ListExistsAsync(_targetSiteUrl, _listTitle);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "CopySelectedViewsAsync — ListExistsAsync failed: {Message}", ex.Message);
                System.Windows.MessageBox.Show(
                    $"Could not reach target site:\n{ex.Message}",
                    "Connection Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                LogAndStatus("View copy cancelled — could not reach target.");
                return;
            }

            if (!listExists)
            {
                System.Windows.MessageBox.Show(
                    $"List \"{_listTitle}\" no longer exists on the target site:\n{_targetSiteUrl}\n\n" +
                    "It may have been deleted. Create it first using the left panel menu.",
                    "List Not Found",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                // Refresh Views tab status to reflect current state
                _ = LoadViewsStatusAsync();
                return;
            }

            // ── Live check 3: field schema still matches ─────────────────────
            bool schemaMatch;
            try
            {
                var targetListId = await _spService.GetListIdByTitleAsync(_targetSiteUrl, _listTitle);
                var targetFields = await _spService.GetListFieldsAsync(_targetSiteUrl, targetListId.ToString());

                var sourceNames = Fields.Select(f => f.InternalName)
                                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                var targetNames = targetFields
                    .Where(f =>
                        (f.InternalName.StartsWith("_x") || !f.InternalName.StartsWith("_")) &&
                        f.TypeAsString != "Computed" &&
                        f.InternalName != "ContentTypeId" &&
                        f.InternalName != "Attachments")
                    .Select(f => f.InternalName)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                schemaMatch = sourceNames.IsSubsetOf(targetNames);

                if (!schemaMatch)
                {
                    var missingFields = sourceNames.Except(targetNames).ToList();
                    _log.Warning("CopySelectedViewsAsync — live schema check failed, missing: {Fields}",
                        string.Join(", ", missingFields));
                }
            }
            catch (Exception ex)
            {
                _log.Error(ex, "CopySelectedViewsAsync — schema check failed: {Message}", ex.Message);
                System.Windows.MessageBox.Show(
                    $"Could not verify field schema on target:\n{ex.Message}",
                    "Verification Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                LogAndStatus("View copy cancelled — schema verification failed.");
                return;
            }

            if (!schemaMatch)
            {
                System.Windows.MessageBox.Show(
                    $"The field schema of list \"{_listTitle}\" has changed on the target site.\n\n" +
                    "View CAML queries reference field names that may not exist on the target —\n" +
                    "copying views could produce broken or empty results.\n\n" +
                    "Run the Compare function from the left panel menu to see\n" +
                    "exactly which fields are missing, then copy the list structure first.",
                    "Schema Mismatch — Copy Unavailable",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                _ = LoadViewsStatusAsync();
                return;
            }

            // ── Select views to copy ─────────────────────────────────────────
            var selectedViews = Views.Where(v => v.IsSelected && !v.ExistsOnTarget).ToList();

            if (selectedViews.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "No views selected.\nCheck at least one missing view.",
                    "Nothing Selected",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
                return;
            }

            // ── Confirm ──────────────────────────────────────────────────────
            var titles  = string.Join(", ", selectedViews.Select(v => v.Title));
            var confirm = System.Windows.MessageBox.Show(
                $"{selectedViews.Count} view(s) will be copied to\n" +
                $"list \"{_listTitle}\" on:\n{_targetSiteUrl}\n\n" +
                $"Views: {titles}\n\n" +
                "Continue?",
                "Confirm Copy Views",
                System.Windows.MessageBoxButton.OKCancel,
                System.Windows.MessageBoxImage.Question);

            if (confirm != System.Windows.MessageBoxResult.OK) return;

            LogAndStatus($"Copying {selectedViews.Count} view(s)...");
            _log.Information("CopySelectedViewsAsync start — views: [{Titles}]", titles);

            try
            {
                await _spService.CopyMissingViewsAsync(
                    _targetSiteUrl,
                    _listTitle,
                    selectedViews);

                // Refresh Views tab — copied views now show as ExistsOnTarget
                await LoadViewsStatusAsync();

                System.Windows.MessageBox.Show(
                    $"{selectedViews.Count} view(s) successfully copied to:\n{_targetSiteUrl}\n\n" +
                    "Refresh the right panel to see the changes.",
                    "Copy Complete",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);

                _log.Information("CopySelectedViewsAsync complete — {Count} views", selectedViews.Count);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR in CopySelectedViewsAsync: {ExType} — {Message}",
                    ex.GetType().Name, ex.Message);
                System.Windows.MessageBox.Show(
                    $"Copy error:\n{ex.Message}",
                    "Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                LogAndStatus($"View copy error: {ex.Message}");
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

            // 2026-06-09: reset per-list flags so Views tab re-checks for a new list
            _targetListExists    = false;
            _targetSchemaMatch   = false;
            _targetViewTitles.Clear();

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
