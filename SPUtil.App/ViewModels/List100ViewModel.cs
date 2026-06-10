using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Services;
using SPUtil.Infrastructure;
using SPUtil.App.Views;
using SPUtil.Views;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Text;
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

        // 2026-06-10: all pre-flight checks moved here — single place, live network calls.
        // No background status loading, no cached flags, no checkbox guards.
        private async Task CopySelectedViewsAsync()
        {
            // 2026-06-10: all checks are live — cached flags are for UX only.
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
                return;
            }

            // ── Live check 3: schema + load target views ─────────────────────
            // We load target views here (not just schema) so we can split
            // selectedViews into toCreate vs toUpdate for the confirmation dialog,
            // and to detect default views that must not be overwritten.
            bool schemaMatch;
            HashSet<string> targetViewTitles;
            List<SPViewData> targetViewsData;
            List<SPFieldData> targetFieldsList;
            try
            {
                var targetListId = await _spService.GetListIdByTitleAsync(_targetSiteUrl, _listTitle);
                targetFieldsList = await _spService.GetListFieldsAsync(_targetSiteUrl, targetListId.ToString());

                // Load full target view data — needed for DefaultView check and create/update split
                targetViewsData = await _spService.GetListViewsAsync(_targetSiteUrl, _listTitle);
                targetViewTitles = targetViewsData
                    .Select(v => v.Title)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);

                _log.Information(
                    "CopySelectedViewsAsync — target views loaded: [{Views}]",
                    string.Join(", ", targetViewsData.Select(v =>
                        $"{v.Title}{(v.DefaultView ? " (default)" : "")}")));

                // Schema check is deferred until after selectedViews is known —
                // we check only the fields actually used in selected views, not all list fields.
                // This avoids false positives where list fields differ in InternalName
                // (e.g. csyk vs filesNumberAlternative) but are not referenced by any view.
                schemaMatch = true; // will be re-evaluated below after selectedViews
            }
            catch (Exception ex)
            {
                _log.Error(ex, "CopySelectedViewsAsync — schema/view check failed: {Message}", ex.Message);
                System.Windows.MessageBox.Show(
                    $"Could not verify target list state:\n{ex.Message}",
                    "Verification Error",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Error);
                LogAndStatus("View copy cancelled — verification failed.");
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
                return;
            }

            // ── Select views ─────────────────────────────────────────────────
            var selectedViews = Views.Where(v => v.IsSelected).ToList();

            if (selectedViews.Count == 0)
            {
                System.Windows.MessageBox.Show(
                    "No views selected.\nCheck at least one view in the list.",
                    "Nothing Selected",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);
                return;
            }

            // ── Schema check — only fields used in selected views ─────────────
            // 2026-06-10 fix: check only ViewFields of selected views, not all list
            // fields. Avoids false positives where list fields have different
            // InternalNames (e.g. 'csyk' vs 'filesNumberAlternative') but are not
            // referenced by any view CAML.
            // NOTE: GetListFieldsAsync filters out Computed fields (LinkTitle, Edit,
            // DocIcon etc.) but these are always present on every SharePoint list.
            // We add them explicitly so they never cause a false mismatch.
            var targetFieldNames = targetFieldsList
                .Select(f => f.InternalName)
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // System Computed fields always present on every SP list — never block on these
            var alwaysPresentFields = new[]
            {
                "LinkTitle", "LinkTitleNoMenu", "LinkTitle2",
                "Edit", "DocIcon", "SelectTitle",
                "LinkFilename", "LinkFilenameNoMenu", "LinkFilename2",
                "ServerUrl", "EncodedAbsUrl", "BaseName",
                "PermMask", "HTML_x0020_File_x0020_Type",
                "_EditMenuTableStart", "_EditMenuTableStart2", "_EditMenuTableEnd",
                "ContentType", "FSObjType", "SortBehavior"
            };
            foreach (var f in alwaysPresentFields)
                targetFieldNames.Add(f);

            var viewFieldsUsed = selectedViews
                .Where(v => v.ViewFields != null)
                .SelectMany(v => v.ViewFields!)
                .Where(f => !string.IsNullOrWhiteSpace(f))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);

            var missingInViews = viewFieldsUsed
                .Where(f => !targetFieldNames.Contains(f))
                .ToList();

            schemaMatch = missingInViews.Count == 0;

            if (!schemaMatch)
            {
                _log.Warning(
                    "CopySelectedViewsAsync — view fields missing on target: [{Fields}]",
                    string.Join(", ", missingInViews));

                System.Windows.MessageBox.Show(
                    $"The following fields are used in the selected view(s)\n" +
                    $"but do not exist on the target list:\n\n" +
                    $"{string.Join("\n", missingInViews.Select(f => $"  • {f}"))}\n\n" +
                    "Copying these views would produce broken or empty results.\n\n" +
                    "Run the Compare function from the left panel menu,\n" +
                    "then copy the list structure first.",
                    "View Fields Missing on Target",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            _log.Information(
                "CopySelectedViewsAsync — schema OK, view fields checked: [{Fields}]",
                string.Join(", ", viewFieldsUsed));

            // ── Split: new views vs existing views to overwrite ───────────────
            var toCreate = selectedViews
                .Where(v => !targetViewTitles.Contains(v.Title))
                .ToList();
            var toUpdate = selectedViews
                .Where(v => targetViewTitles.Contains(v.Title))
                .ToList();

            _log.Information(
                "CopySelectedViewsAsync — to create: [{Create}] | to update: [{Update}]",
                string.Join(", ", toCreate.Select(v => v.Title)),
                string.Join(", ", toUpdate.Select(v => v.Title)));

            // ── Block overwrite of default view(s) on target ─────────────────
            // SharePoint requires exactly one default view per list.
            // Overwriting the default view risks leaving the list in a broken state
            // if anything fails mid-operation. Instruct user to do it via SP UI.
            var blocked = toUpdate
                .Where(v => targetViewsData
                    .Any(t => t.Title.Equals(v.Title, StringComparison.OrdinalIgnoreCase)
                           && t.DefaultView))
                .ToList();

            if (blocked.Count > 0)
            {
                var blockedNames = string.Join("\n", blocked.Select(v => $"  • {v.Title}"));
                _log.Warning(
                    "CopySelectedViewsAsync — blocked: default view(s) on target cannot be overwritten: [{Views}]",
                    string.Join(", ", blocked.Select(v => v.Title)));

                System.Windows.MessageBox.Show(
                    $"The following view(s) are the DEFAULT view on the target list\n" +
                    $"and cannot be overwritten by this tool:\n\n" +
                    $"{blockedNames}\n\n" +
                    "Overwriting the default view risks breaking the list if\n" +
                    "anything fails mid-operation.\n\n" +
                    "To modify the default view, use the SharePoint list settings\n" +
                    "in the browser directly.",
                    "Default View — Cannot Overwrite",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Warning);
                return;
            }

            // ── Build confirmation message ────────────────────────────────────
            var sb = new StringBuilder();

            if (toCreate.Count > 0)
            {
                sb.AppendLine($"NEW — will be created ({toCreate.Count}):");
                foreach (var v in toCreate)
                    sb.AppendLine($"  + {v.Title}");
                sb.AppendLine();
            }

            if (toUpdate.Count > 0)
            {
                sb.AppendLine($"EXISTING — will be OVERWRITTEN ({toUpdate.Count}):");
                foreach (var v in toUpdate)
                    sb.AppendLine($"  ⚠ {v.Title}");
                sb.AppendLine();
            }

            sb.AppendLine($"Target list: \"{_listTitle}\"");
            sb.AppendLine($"Target site: {_targetSiteUrl}");
            sb.AppendLine();
            sb.Append("Continue?");

            var confirm = System.Windows.MessageBox.Show(
                sb.ToString(),
                "Confirm Copy Views",
                System.Windows.MessageBoxButton.OKCancel,
                toUpdate.Count > 0
                    ? System.Windows.MessageBoxImage.Warning
                    : System.Windows.MessageBoxImage.Question);

            if (confirm != System.Windows.MessageBoxResult.OK) return;

            LogAndStatus($"Copying {selectedViews.Count} view(s) ({toCreate.Count} new, {toUpdate.Count} overwrite)...");
            _log.Information(
                "CopySelectedViewsAsync start — create: [{Create}] | overwrite: [{Update}]",
                string.Join(", ", toCreate.Select(v => v.Title)),
                string.Join(", ", toUpdate.Select(v => v.Title)));

            try
            {
                await _spService.CopyMissingViewsAsync(
                    _targetSiteUrl,
                    _listTitle,
                    selectedViews);

                var resultMsg = new StringBuilder();
                if (toCreate.Count > 0)
                    resultMsg.AppendLine($"Created: {string.Join(", ", toCreate.Select(v => v.Title))}");
                if (toUpdate.Count > 0)
                    resultMsg.AppendLine($"Overwritten: {string.Join(", ", toUpdate.Select(v => v.Title))}");
                resultMsg.AppendLine();
                resultMsg.Append("Refresh the right panel to see the changes.");

                System.Windows.MessageBox.Show(
                    resultMsg.ToString(),
                    "Copy Complete",
                    System.Windows.MessageBoxButton.OK,
                    System.Windows.MessageBoxImage.Information);

                LogAndStatus($"Done. {toCreate.Count} created, {toUpdate.Count} overwritten.");
                _log.Information(
                    "CopySelectedViewsAsync complete — created: {Created}, overwritten: {Updated}",
                    toCreate.Count, toUpdate.Count);
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
