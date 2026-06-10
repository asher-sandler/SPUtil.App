using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Infrastructure;
using SPUtil.Services;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using Serilog;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace SPUtil.App.ViewModels
{
    public class PagesViewModel : BindableBase
    {
        private static readonly ILogger _log = Log.ForContext<PagesViewModel>();

        private readonly ISharePointService _spService;
        private string  _siteUrl       = string.Empty;
        private string  _targetSiteUrl = string.Empty;
        private string  _statusMessage = "Ready";
        private ObservableCollection<SPFileData>    _pages    = new();
        private ObservableCollection<SPWebPartData> _webParts = new();
        private SPFileData?    _selectedPage;
        private SPWebPartData? _selectedWebPart;

        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        public ObservableCollection<SPFileData> Pages
        {
            get => _pages;
            set => SetProperty(ref _pages, value);
        }

        public ObservableCollection<SPWebPartData> WebParts
        {
            get => _webParts;
            set => SetProperty(ref _webParts, value);
        }

        /// <summary>Currently selected WebPart in the bottom grid</summary>
        public SPWebPartData? SelectedWebPart
        {
            get => _selectedWebPart;
            set => SetProperty(ref _selectedWebPart, value);
        }

        private bool _isSourceMode;
        public bool IsSourceMode
        {
            get => _isSourceMode;
            set => SetProperty(ref _isSourceMode, value);
        }

        public SPFileData? SelectedPage
        {
            get => _selectedPage;
            set
            {
                if (SetProperty(ref _selectedPage, value) && value != null)
                    _ = LoadWebPartsAsync(value.FullPath);
            }
        }

        // ── Commands ──────────────────────────────────────────────────────────
        public DelegateCommand GetAllPropertiesCommand    { get; }
        public DelegateCommand ShowWebPartsPreviewCommand { get; }
        public DelegateCommand CopyPageCommand            { get; }
        public DelegateCommand DeletePageCommand          { get; }
        public DelegateCommand RenamePageCommand          { get; }
        public DelegateCommand ComparePageCommand         { get; }
        public DelegateCommand SyncPropertiesCommand      { get; }

        // ── Per-WebPart commands (operate on SelectedWebPart) ─────────────────
        /// <summary>Compare selected WebPart properties with the same-titled WP on target page</summary>
        public DelegateCommand CompareWebPartCommand         { get; }
        /// <summary>Copy properties of selected WebPart to the same-titled WP on target page</summary>
        public DelegateCommand CopyWebPartPropertiesCommand  { get; }
        /// <summary>Copy all properties of selected WebPart to clipboard (with page/WP header)</summary>
        public DelegateCommand CopyWpToClipboardCommand      { get; }
        /// <summary>Generate PowerShell script with all WebParts as embedded JSON, show in preview</summary>
        public DelegateCommand ExportWpToPowerShellCommand   { get; }

        public PagesViewModel(ISharePointService spService)
        {
            _spService = spService;

            GetAllPropertiesCommand = new DelegateCommand(() =>
            {
                Debug.WriteLine($">>> [Pages] {WebParts.Count} web parts for '{SelectedPage?.Name}'");
                foreach (var wp in WebParts)
                {
                    Debug.WriteLine($"WebPart: {wp.Title} ({wp.StorageKey})");
                    foreach (var prop in wp.Properties)
                        Debug.WriteLine($"   {prop.Key}: {prop.Value}");
                }
                StatusMessage = "Property data written to Output";
            });

            ShowWebPartsPreviewCommand = new DelegateCommand(() =>
            {
                if (SelectedPage == null)
                {
                    MessageBox.Show("Select a page from the list above.",
                        "No page selected", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (!WebParts.Any())
                {
                    MessageBox.Show("The selected page does not contain any web parts or they have not yet loaded.",
                        "No web parts", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                var win = new SPUtil.App.Views.UniversalPreviewWindow
                {
                    Title  = $"WebParts — {SelectedPage.Name}",
                    Owner  = Application.Current.MainWindow,
                    Width  = 1000,
                    Height = 680
                };
                var vm = new WebPartsPreviewViewModel(WebParts, SelectedPage.Name, win);
                win.DataContext = vm;
                win.ShowDialog();
            });

            CopyPageCommand   = new DelegateCommand(async () => await ExecuteCopyPageAsync());
            DeletePageCommand = new DelegateCommand(async () => await ExecuteDeletePageAsync());
            RenamePageCommand = new DelegateCommand(async () => await ExecuteRenamePageAsync());
            ComparePageCommand    = new DelegateCommand(async () => await ExecuteComparePageAsync());
            SyncPropertiesCommand = new DelegateCommand(async () => await ExecuteSyncPropertiesAsync());

            CompareWebPartCommand = new DelegateCommand(
                async () => await ExecuteCompareWebPartAsync(),
                () => SelectedWebPart != null && !string.IsNullOrEmpty(_targetSiteUrl))
                .ObservesProperty(() => SelectedWebPart);

            CopyWebPartPropertiesCommand = new DelegateCommand(
                async () => await ExecuteCopyWebPartPropertiesAsync(),
                () => SelectedWebPart != null && !string.IsNullOrEmpty(_targetSiteUrl))
                .ObservesProperty(() => SelectedWebPart);

            CopyWpToClipboardCommand = new DelegateCommand(
                () => ExecuteCopyWpToClipboard(),
                () => SelectedWebPart != null)
                .ObservesProperty(() => SelectedWebPart);

            ExportWpToPowerShellCommand = new DelegateCommand(
                () => ExecuteExportWpToPowerShell(),
                () => WebParts != null && WebParts.Any())
                .ObservesProperty(() => WebParts);
        }

        // ── Called by MainWindowViewModel after creating this VM ──────────────
        public void SetTargetSiteUrl(string url) => _targetSiteUrl = url;


        // ═══════════════════════════════════════════════════════════════════════
        //  Copy Page
        // ═══════════════════════════════════════════════════════════════════════
        private async Task ExecuteCopyPageAsync()
        {
            _log.Debug("CopyPage started. SelectedPage={Page} TargetSite={Site}",
                SelectedPage?.Name, _targetSiteUrl);

            if (SelectedPage == null)
            {
                _log.Warning("CopyPage aborted — no page selected");
                MessageBox.Show("Please select a page to copy.",
                    "No page selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrEmpty(_targetSiteUrl))
            {
                _log.Warning("CopyPage aborted — no target site configured");
                MessageBox.Show("Connect to target site (right panel).",
                    "No target site", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string sourceName    = System.IO.Path.GetFileNameWithoutExtension(SelectedPage.Name);
            string sourceSubfolder = ComputeSubfolderPath(SelectedPage.FullPath);
            string sourceInfo      = $"Page : {SelectedPage.Name}\nPath : {SelectedPage.FullPath}\nSite : {_siteUrl}";

            // ── Step 1: Dialog — name + optional path ─────────────────────────
            var dialog = new SPUtil.Views.CopyPageDialog(
                sourceName, _targetSiteUrl, sourceInfo, sourceSubfolder)
            {
                Owner = Application.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            string targetName    = dialog.TargetPageName;
            string subfolderPath = (dialog.KeepFolderPath && !string.IsNullOrEmpty(sourceSubfolder))
                                   ? sourceSubfolder : string.Empty;

            // ── Step 2: Check if page already exists on target ────────────────
            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = Application.Current.MainWindow
            };
            infoWin.Show();
            infoWin.UpdateMessage("Checking target site...");

            bool exists = await _spService.PageExistsAsync(_targetSiteUrl, targetName);
            _log.Information("PageExists check: {Page} exists={Exists} on {Site}",
                targetName, exists, _targetSiteUrl);

            if (exists)
            {
                // Page exists — ask what to do (Replace or Rename)
                infoWin.Close();

                var existsDialog = MessageBox.Show(
                    $"Page '{targetName}.aspx' already exists on target site.\n\n" +
                    $"Replace — delete existing page and create from source\n" +
                    $"No — rename existing to '{targetName}_old' first, then create",
                    "Page Already Exists",
                    MessageBoxButton.YesNoCancel,
                    MessageBoxImage.Warning);

                if (existsDialog == MessageBoxResult.Cancel) return;

                infoWin = new SPUtil.Views.OperationInfoWindow
                {
                    Owner = Application.Current.MainWindow
                };
                infoWin.Show();

                if (existsDialog == MessageBoxResult.Yes)
                {
                    // Replace — delete existing
                    infoWin.UpdateMessage($"Deleting existing page '{targetName}'...");
                    try
                    {
                        await _spService.DeletePageAsync(_targetSiteUrl, targetName);
                    }
                    catch (Exception ex)
                    {
                        _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                        infoWin.Close();
                        MessageBox.Show($"Error deleting existing page:\n{ex.Message}",
                            "Delete Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
                else
                {
                    // No — rename existing to _old first
                    string oldName = targetName + "_old";
                    infoWin.UpdateMessage($"Renaming '{targetName}' → '{oldName}'...");
                    try
                    {
                        await _spService.RenamePageAsync(_targetSiteUrl, targetName, oldName);
                    }
                    catch (Exception ex)
                    {
                        _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                        infoWin.Close();
                        MessageBox.Show($"Error renaming existing page:\n{ex.Message}",
                            "Rename Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        return;
                    }
                }
            }

            // ── Step 3: Read snapshot from source ─────────────────────────────
            infoWin.UpdateMessage("Reading source page (layout + WebParts)...");
            PageSnapshot snapshot;
            try
            {
                _log.Information("Reading snapshot: {Page} from {Site}", SelectedPage.FullPath, _siteUrl);
                snapshot = await _spService.GetPageSnapshotAsync(_siteUrl, SelectedPage.FullPath);
                _log.Information("Snapshot read OK — {Count} WebPart(s)", snapshot.WebParts.Count);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "Failed to read snapshot: {Page}", SelectedPage.FullPath);
                infoWin.Close();
                MessageBox.Show($"Error reading source page:\n{ex.Message}",
                    "Snapshot Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // ── Step 4: Create page on target (with optional subfolder) ────────
            int wpCount = snapshot.WebParts.Count;
            string pathLabel = string.IsNullOrEmpty(subfolderPath)
                ? "" : $" in Pages/{subfolderPath}";
            infoWin.UpdateMessage($"Creating '{targetName}'{pathLabel} with {wpCount} WebPart(s)...");
            try
            {
                _log.Information("Creating page {Name} in subfolder='{Sub}' on {Site}",
                    targetName, subfolderPath, _targetSiteUrl);
                await _spService.CreatePageFromSnapshotAsync(
                    _targetSiteUrl, targetName, snapshot, subfolderPath);
                _log.Information("Page created successfully: {Name}", targetName);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "Failed to create page {Name} on {Site}", targetName, _targetSiteUrl);
                infoWin.Close();
                MessageBox.Show($"Error creating target page:\n{ex.Message}",
                    "Create Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            infoWin.Close();
            StatusMessage = $"✔ '{targetName}' copied → {_targetSiteUrl}";

            MessageBox.Show(
                $"Page '{targetName}' created successfully on:\n{_targetSiteUrl}\n\n" +
                $"WebParts copied: {wpCount}\n\n" +
                $"⚠ Page permissions must be configured manually on the target site.",
                "Copy Complete",
                MessageBoxButton.OK,
                MessageBoxImage.Information);
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  Delete Page
        // ═══════════════════════════════════════════════════════════════════════
        private async Task ExecuteDeletePageAsync()
        {
            _log.Debug("DeletePage started. Page={Page}", SelectedPage?.Name);
            if (SelectedPage == null)
            {
                MessageBox.Show("Please select a page to delete.",
                    "No page selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var confirm = MessageBox.Show(
                $"Delete page '{SelectedPage.Name}'?\n\nThis cannot be undone.",
                "Confirm Delete", MessageBoxButton.YesNo, MessageBoxImage.Warning);
            if (confirm != MessageBoxResult.Yes) return;

            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = Application.Current.MainWindow
            };
            infoWin.Show();
            infoWin.UpdateMessage($"Deleting '{SelectedPage.Name}'...");

            try
            {
                _log.Information("Deleting page {Page} from {Site}", SelectedPage.Name, _siteUrl);
                await _spService.DeletePageAsync(_siteUrl, SelectedPage.Name);
                _log.Information("Page deleted: {Page}", SelectedPage.Name);

                var removed = Pages.FirstOrDefault(p => p.FullPath == SelectedPage.FullPath);
                if (removed != null) Pages.Remove(removed);

                SelectedPage  = null;
                StatusMessage = "Page deleted.";
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                StatusMessage = $"Delete error: {ex.Message}";
                MessageBox.Show($"Error deleting page:\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { infoWin.Close(); }
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  Rename Page
        // ═══════════════════════════════════════════════════════════════════════
        private async Task ExecuteRenamePageAsync()
        {
            _log.Debug("RenamePage started. Page={Page}", SelectedPage?.Name);
            if (SelectedPage == null)
            {
                MessageBox.Show("Please select a page to rename.",
                    "No page selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string currentName = System.IO.Path.GetFileNameWithoutExtension(SelectedPage.Name);

            var dialog = new SPUtil.Views.CopyPageDialog(
                currentName, _siteUrl,
                $"Rename page: {SelectedPage.Name}\nSite: {_siteUrl}")
            {
                Title = "Rename Page",
                Owner = Application.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            string newName = dialog.TargetPageName;
            if (newName.Equals(currentName, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("New name is the same as current.",
                    "No Change", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = Application.Current.MainWindow
            };
            infoWin.Show();
            infoWin.UpdateMessage($"Renaming '{currentName}' → '{newName}'...");

            try
            {
                await _spService.RenamePageAsync(_siteUrl, currentName, newName);

                // Update local collection
                if (SelectedPage != null)
                {
                    SelectedPage.Name     = newName + ".aspx";
                    SelectedPage.FullPath = SelectedPage.FullPath.Replace(
                        currentName + ".aspx", newName + ".aspx",
                        StringComparison.OrdinalIgnoreCase);
                    // Force grid refresh
                    var tmp = new ObservableCollection<SPFileData>(Pages);
                    Pages = tmp;
                }

                StatusMessage = $"Renamed: {currentName} → {newName}";
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                StatusMessage = $"Rename error: {ex.Message}";
                MessageBox.Show($"Error renaming page:\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally { infoWin.Close(); }
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  Compare Page
        // ═══════════════════════════════════════════════════════════════════════
        private async Task ExecuteComparePageAsync()
        {
            _log.Debug("ComparePage started. Source={Page} Site={Site}",
                SelectedPage?.Name, _siteUrl);
            if (SelectedPage == null)
            {
                MessageBox.Show("Please select a page to compare.",
                    "No page selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrEmpty(_targetSiteUrl))
            {
                MessageBox.Show("Connect to target site (right panel).",
                    "No target site", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string sourceName = System.IO.Path.GetFileNameWithoutExtension(SelectedPage.Name);
            var dialog = new SPUtil.Views.CopyPageDialog(
                sourceName, _targetSiteUrl,
                $"Compare: {SelectedPage.FullPath}\nSource site: {_siteUrl}")
            {
                Title = "Compare Page — enter target page name",
                Owner = Application.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            string targetPageName = dialog.TargetPageName;

            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = Application.Current.MainWindow
            };
            infoWin.Show();

            PageCompareResult compareResult;
            string formattedText;

            try
            {
                infoWin.UpdateMessage("Reading source page snapshot...");
                var sourceSnapshot = await _spService.GetPageSnapshotAsync(
                    _siteUrl, SelectedPage.FullPath);

                infoWin.UpdateMessage("Reading target page snapshot...");
                string targetRelUrl = await _spService.GetPageRelativeUrlAsync(
                    _targetSiteUrl, targetPageName);
                var targetSnapshot = await _spService.GetPageSnapshotAsync(
                    _targetSiteUrl, targetRelUrl);

                infoWin.UpdateMessage("Comparing...");
                compareResult = await _spService.ComparePageSnapshotsStructured(
                    sourceSnapshot, targetSnapshot,
                    sourceSiteUrl: _siteUrl,
                    targetSiteUrl: _targetSiteUrl);

                // Store target URL for InsertPlaceholders button
                compareResult.TargetUrl = targetRelUrl;

                formattedText = _spService.FormatCompareResult(compareResult);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                infoWin.Close();
                MessageBox.Show($"Error during comparison:\n{ex.Message}",
                    "Compare Error", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            infoWin.Close();

            var win = new SPUtil.App.Views.UniversalPreviewWindow
            {
                Title  = $"Compare: {SelectedPage.Name}  ↔  {targetPageName}.aspx",
                Owner  = Application.Current.MainWindow,
                Width  = 1050,
                Height = 720
            };

            var vm = new PageCompareViewModel(
                compareResult, formattedText, win, _spService,
                onPlaceholdersInserted: () =>
                {
                    // Refresh status after placeholders inserted
                    StatusMessage = $"✔ Placeholders inserted on target — " +
                                    $"add WebParts manually then run Sync Properties";
                });

            win.DataContext = vm;
            win.ShowDialog();

            StatusMessage = compareResult.IsIdentical
                ? "✔ Pages are identical"
                : $"Differences: {compareResult.ModifiedCount} modified, " +
                  $"{compareResult.AddedCount} added, {compareResult.RemovedCount} removed";
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  Sync Properties
        //  Finds SPUTIL placeholders on the selected page, reads WebPart settings
        //  from source, applies them to matching WebParts, removes placeholders.
        // ═══════════════════════════════════════════════════════════════════════
        private async Task ExecuteSyncPropertiesAsync()
        {
            _log.Debug("SyncProperties started. Page={Page}", SelectedPage?.Name);
            if (SelectedPage == null)
            {
                MessageBox.Show("Please select a page to sync.",
                    "No page selected", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // Confirm
            var confirm = MessageBox.Show(
                $"Sync Properties will:\n" +
                $"1. Find all [WebPart Placeholder] blocks on '{SelectedPage.Name}'\n" +
                $"2. Copy settings from the source WebParts\n" +
                $"3. Apply settings to matching WebParts on this page\n" +
                $"4. Remove the placeholder blocks\n\n" +
                $"Continue?",
                "Sync Properties", MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;

            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = Application.Current.MainWindow
            };
            infoWin.Show();
            infoWin.UpdateMessage("Checking for placeholders...");

            try
            {
                bool hasPlaceholders = await _spService.PageHasPlaceholdersAsync(
                    _siteUrl, SelectedPage.FullPath);

                if (!hasPlaceholders)
                {
                    infoWin.Close();
                    MessageBox.Show(
                        "No WebPart Placeholders found on this page.\n" +
                        "Run Compare Page first, then Insert Placeholders.",
                        "No Placeholders", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                infoWin.UpdateMessage("Syncing WebPart properties...");
                var result = await _spService.SyncPropertiesAsync(
                    _siteUrl, SelectedPage.FullPath);

                infoWin.Close();

                string message = $"Sync complete.\n\n{result.Summary}";
                if (result.Errors.Any())
                {
                    message += "\n\nDetails:\n" + string.Join("\n", result.Errors);
                }

                MessageBox.Show(message,
                    result.IsSuccess ? "Sync Complete" : "Sync Complete (with warnings)",
                    MessageBoxButton.OK,
                    result.IsSuccess ? MessageBoxImage.Information : MessageBoxImage.Warning);

                StatusMessage = $"Sync: {result.Summary}";

                // Reload WebParts to reflect changes
                await LoadWebPartsAsync(SelectedPage.FullPath);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                infoWin.Close();
                StatusMessage = $"Sync error: {ex.Message}";
                MessageBox.Show($"Error during sync:\n{ex.Message}",
                    "Sync Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  Compare WebPart — shows diff of one WP between source and target page
        // ═══════════════════════════════════════════════════════════════════════
        private async Task ExecuteCompareWebPartAsync()
        {
            if (SelectedWebPart == null || SelectedPage == null) return;

            // Ask which page on target to compare with
            string sourceName = System.IO.Path.GetFileNameWithoutExtension(SelectedPage.Name);
            var dialog = new SPUtil.Views.ComparePageDialog(
                sourceName, _targetSiteUrl,
                $"WebPart  : {SelectedWebPart.Title}\nSource   : {SelectedPage.FullPath}\nSite     : {_siteUrl}")
            {
                Title = "Compare WebPart — enter target page name",
                Owner = Application.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            string targetPageName = dialog.TargetPageName;

            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = Application.Current.MainWindow
            };
            infoWin.Show();
            infoWin.UpdateMessage("Reading target page snapshot...");

            try
            {
                string targetRelUrl = await _spService.GetPageRelativeUrlAsync(
                    _targetSiteUrl, targetPageName);
                var targetSnapshot = await _spService.GetPageSnapshotAsync(
                    _targetSiteUrl, targetRelUrl);

                // Find matching WebPart on target by Title
                var targetWp = targetSnapshot.WebParts
                    .OrderBy(w => Math.Abs(w.VisualPosition - SelectedWebPart.VisualPosition))
                    .FirstOrDefault(w => string.Equals(w.Title, SelectedWebPart.Title,
                        StringComparison.OrdinalIgnoreCase));

                infoWin.Close();

                if (targetWp == null)
                {
                    MessageBox.Show(
                        $"WebPart '{SelectedWebPart.Title}' not found on target page '{targetPageName}'.\n\n" +
                        $"Use Copy WebPart Properties to copy its settings after adding it manually.",
                        "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // Build diff text for this single WebPart
                var sb = new System.Text.StringBuilder();
                sb.AppendLine($"=== WebPart Comparison: {SelectedWebPart.Title} ===");
                sb.AppendLine($"Source : {_siteUrl}{SelectedPage.FullPath}");
                sb.AppendLine($"Target : {_targetSiteUrl}{targetRelUrl}");
                sb.AppendLine(new string('═', 70));

                var skipProps = new System.Collections.Generic.HashSet<string>(
                    System.StringComparer.OrdinalIgnoreCase)
                {
                    "AllowClose","AllowConnect","AllowEdit","AllowHide","AllowMinimize",
                    "AllowZoneChange","AuthorizationFilter","CatalogIconImageUrl",
                    "ChromeState","ChromeType","Direction","ExportMode","HelpMode",
                    "HelpUrl","Hidden","ImportErrorMessage","TitleIconImageUrl","TitleUrl"
                };

                int diffCount = 0;
                var allKeys = SelectedWebPart.Properties.Keys
                    .Union(targetWp.Properties.Keys, System.StringComparer.OrdinalIgnoreCase)
                    .Where(k => !skipProps.Contains(k))
                    .OrderBy(k => k);

                foreach (var key in allKeys)
                {
                    SelectedWebPart.Properties.TryGetValue(key, out var sv); sv ??= "";
                    targetWp.Properties.TryGetValue(key, out var tv);        tv ??= "";

                    if (!string.Equals(sv, tv, System.StringComparison.Ordinal))
                    {
                        diffCount++;
                        sb.AppendLine();
                        sb.AppendLine($"✏ {key}");
                        sb.AppendLine($"  source: {(string.IsNullOrEmpty(sv) ? "(empty)" : sv)}");
                        sb.AppendLine($"  target: {(string.IsNullOrEmpty(tv) ? "(empty)" : tv)}");
                    }
                }

                if (diffCount == 0)
                    sb.AppendLine("\n✔ WebPart properties are IDENTICAL.");
                else
                    sb.AppendLine($"\nTotal differences: {diffCount}");

                // Show in preview window
                var win = new SPUtil.App.Views.UniversalPreviewWindow
                {
                    Title  = $"WP Compare: {SelectedWebPart.Title}",
                    Owner  = Application.Current.MainWindow,
                    Width  = 900,
                    Height = 620
                };
                var vm = new WebPartsPreviewViewModel(
                    new System.Collections.Generic.List<SPWebPartData>(),
                    SelectedWebPart.Title, win);
                // Override PreviewText directly
                vm.PreviewText = sb.ToString();
                win.DataContext = vm;
                win.ShowDialog();
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                infoWin.Close();
                MessageBox.Show($"Compare error:\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  Copy WebPart Properties — copies settings of one WP to matching WP
        //  on the target page (by Title + closest position)
        // ═══════════════════════════════════════════════════════════════════════
        private async Task ExecuteCopyWebPartPropertiesAsync()
        {
            if (SelectedWebPart == null || SelectedPage == null) return;

            // Ask target page name
            string sourceName = System.IO.Path.GetFileNameWithoutExtension(SelectedPage.Name);
            var dialog = new SPUtil.Views.ComparePageDialog(
                sourceName, _targetSiteUrl,
                $"WebPart  : {SelectedWebPart.Title}\nSource   : {SelectedPage.FullPath}\nSite     : {_siteUrl}")
            {
                Title = "Copy WebPart Properties — enter target page name",
                Owner = Application.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            string targetPageName = dialog.TargetPageName;

            var confirm = MessageBox.Show(
                $"Copy all properties of\n'{SelectedWebPart.Title}'\n" +
                $"from source to the matching WebPart on '{targetPageName}'?\n\n" +
                $"Target site: {_targetSiteUrl}",
                "Confirm Copy Properties",
                MessageBoxButton.YesNo, MessageBoxImage.Question);
            if (confirm != MessageBoxResult.Yes) return;

            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = Application.Current.MainWindow
            };
            infoWin.Show();
            infoWin.UpdateMessage("Reading target page...");

            try
            {
                string targetRelUrl = await _spService.GetPageRelativeUrlAsync(
                    _targetSiteUrl, targetPageName);
                var targetSnapshot = await _spService.GetPageSnapshotAsync(
                    _targetSiteUrl, targetRelUrl);

                // Find best-matching WebPart on target
                var targetWp = targetSnapshot.WebParts
                    .OrderBy(w => Math.Abs(w.VisualPosition - SelectedWebPart.VisualPosition))
                    .FirstOrDefault(w => string.Equals(w.Title, SelectedWebPart.Title,
                        StringComparison.OrdinalIgnoreCase));

                if (targetWp == null)
                {
                    infoWin.Close();
                    MessageBox.Show(
                        $"WebPart '{SelectedWebPart.Title}' not found on '{targetPageName}'.\n" +
                        $"Add it to the target page first.",
                        "Not Found", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                infoWin.UpdateMessage($"Copying properties to '{targetWp.Title}'...");

                // Get ExportXml from source to extract all properties
                string exportXml = await _spService.GetWebPartExportXmlAsync(
                    _siteUrl, SelectedPage.FullPath, SelectedWebPart.StorageKey);

                if (string.IsNullOrEmpty(exportXml))
                {
                    infoWin.Close();
                    MessageBox.Show("Could not download source WebPart settings.",
                        "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                // Apply to target
                var request = new WebPartUpdateRequest
                {
                    StorageKey         = targetWp.StorageKey,
                    PropertiesToUpdate = ParseExportXmlProperties(exportXml)
                };

                await _spService.UpdateWebPartAsync(_targetSiteUrl, targetRelUrl, request);

                infoWin.Close();
                StatusMessage = $"✔ Properties copied to '{targetWp.Title}' on '{targetPageName}'";
                MessageBox.Show(
                    $"Properties of '{SelectedWebPart.Title}' copied successfully\nto '{targetPageName}' on target site.",
                    "Done", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                infoWin.Close();
                MessageBox.Show($"Error copying properties:\n{ex.Message}",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>Parses custom properties from exportwp.aspx XML, skipping system ones.</summary>
        private static System.Collections.Generic.Dictionary<string, string> ParseExportXmlProperties(string xml)
        {
            var result = new System.Collections.Generic.Dictionary<string, string>(
                StringComparer.OrdinalIgnoreCase);
            var skip = new System.Collections.Generic.HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "AllowClose","AllowConnect","AllowEdit","AllowHide","AllowMinimize",
                "AllowZoneChange","AuthorizationFilter","CatalogIconImageUrl",
                "ChromeState","ChromeType","Direction","ExportMode","HelpMode",
                "HelpUrl","Hidden","ImportErrorMessage","TitleIconImageUrl","TitleUrl",
                "Title","Description"
            };
            try
            {
                var doc = System.Xml.Linq.XDocument.Parse(xml);
                foreach (var prop in doc.Descendants().Where(e => e.Name.LocalName == "property"))
                {
                    string name = prop.Attribute("name")?.Value ?? "";
                    if (!string.IsNullOrEmpty(name) && !skip.Contains(name))
                        result[name] = prop.Value ?? "";
                }
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                System.Diagnostics.Debug.WriteLine($"[ParseExportXml] {ex.Message}");
            }
            return result;
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  Copy WebPart properties to clipboard
        //  Format: page name + WP title + all key:value properties
        // ═══════════════════════════════════════════════════════════════════════
        private void ExecuteCopyWpToClipboard()
        {
            if (SelectedWebPart == null) return;

            var sb = new System.Text.StringBuilder();

            // Header — page and WebPart identity
            sb.AppendLine($"Page     : {SelectedPage?.Name ?? _siteUrl}");
            sb.AppendLine($"Site     : {_siteUrl}");
            sb.AppendLine($"WebPart  : {SelectedWebPart.Title}");
            sb.AppendLine($"Zone     : {SelectedWebPart.ZoneId}");
            sb.AppendLine($"Position : {SelectedWebPart.VisualPosition}");
            sb.AppendLine($"StorageKey: {SelectedWebPart.StorageKey}");
            sb.AppendLine(new string('─', 60));

            // All properties
            foreach (var kv in SelectedWebPart.Properties.OrderBy(k => k.Key))
            {
                string val = kv.Value ?? "";
                // Indent multi-line values
                if (val.Contains('\n') || val.Contains('\r'))
                    val = "\n" + string.Join("\n",
                        val.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries)
                           .Select(line => "              " + line));

                sb.AppendLine($"{kv.Key,-30}: {val}");
            }

            try
            {
                System.Windows.Clipboard.SetText(sb.ToString());
                StatusMessage = $"✔ Copied to clipboard: {SelectedWebPart.Title} ({SelectedWebPart.Properties.Count} properties)";
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                StatusMessage = $"Clipboard error: {ex.Message}";
            }
        }

        /// <summary>
        /// Extracts the subfolder path within Pages from a full server-relative page URL.
        /// /home/Agriculture/Pages/Dean/Candidate.aspx → "Dean"
        /// /home/Agriculture/Pages/FacultyAdmin/Sub/Page.aspx → "FacultyAdmin/Sub"
        /// /home/Agriculture/Pages/Page.aspx → ""
        /// </summary>
        private static string ComputeSubfolderPath(string pageRelativeUrl)
        {
            if (string.IsNullOrEmpty(pageRelativeUrl)) return string.Empty;
            string url = pageRelativeUrl.Replace('\\', '/');
            int pagesIdx = url.IndexOf("/pages/", StringComparison.OrdinalIgnoreCase);
            if (pagesIdx < 0) return string.Empty;
            string afterPages = url.Substring(pagesIdx + "/pages/".Length);
            int lastSlash = afterPages.LastIndexOf('/');
            return lastSlash <= 0 ? string.Empty : afterPages.Substring(0, lastSlash);
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  Export WebParts as PowerShell script
        //  Generates a .ps1 file that:
        //    1. Contains embedded JSON (as a here-string @"..."@)
        //    2. Defines $PageSnapshot as a PowerShell hashtable
        //    3. Outputs page name, WebPart names, properties and values on run
        // ═══════════════════════════════════════════════════════════════════════
        private void ExecuteExportWpToPowerShell()
        {
            if (!WebParts.Any() || SelectedPage == null) return;

            string pageName      = SelectedPage.Name;
            string pageNameNoExt = System.IO.Path.GetFileNameWithoutExtension(pageName);
            string now           = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            string subfolder     = ComputeSubfolderPath(SelectedPage.FullPath);

            // ── Build JSON (single-quoted here-string — $ signs won't be expanded) ──
            var jsonSb = new System.Text.StringBuilder();
            jsonSb.AppendLine("{");
            jsonSb.AppendLine($"  \"Page\": \"{EscapeJson(pageName)}\",");
            jsonSb.AppendLine($"  \"Site\": \"{EscapeJson(_targetSiteUrl)}\",");
            jsonSb.AppendLine($"  \"Path\": \"{EscapeJson(SelectedPage.FullPath)}\",");
            jsonSb.AppendLine($"  \"Subfolder\": \"{EscapeJson(subfolder)}\",");
            jsonSb.AppendLine($"  \"Exported\": \"{now}\",");
            jsonSb.AppendLine("  \"WebParts\": [");

            var wpList = WebParts.ToList();
            for (int i = 0; i < wpList.Count; i++)
            {
                var  wp     = wpList[i];
                bool lastWp = i == wpList.Count - 1;
                jsonSb.AppendLine("    {");
                jsonSb.AppendLine($"      \"Title\": \"{EscapeJson(wp.Title)}\",");
                jsonSb.AppendLine($"      \"Position\": {wp.VisualPosition},");
                jsonSb.AppendLine($"      \"Zone\": \"{EscapeJson(wp.ZoneId)}\",");
                jsonSb.AppendLine($"      \"StorageKey\": \"{wp.StorageKey}\",");
                jsonSb.AppendLine("      \"Properties\": {");
                var props = wp.Properties.OrderBy(k => k.Key).ToList();
                for (int j = 0; j < props.Count; j++)
                {
                    bool   lastProp = j == props.Count - 1;
                    string val      = EscapeJson(props[j].Value ?? "");
                    string comma    = lastProp ? "" : ",";
                    jsonSb.AppendLine($"        \"{EscapeJson(props[j].Key)}\": \"{val}\"{comma}");
                }
                jsonSb.AppendLine("      }");
                jsonSb.AppendLine(lastWp ? "    }" : "    },");
            }
            jsonSb.AppendLine("  ]");
            jsonSb.Append("}");
            string json = jsonSb.ToString();

            // ── Normalize source URL (remove trailing digit from hostname) ────────

            string normalizedSite; // = _siteUrl.TrimEnd('/');
            string targetSite;
            if (string.IsNullOrWhiteSpace(_targetSiteUrl))
            {
                targetSite = _siteUrl.TrimEnd('/');
            }
            else
            {
                targetSite = _targetSiteUrl.TrimEnd('/');
            }
            try
            {
                var uri    = new Uri(_siteUrl);
                var parts  = uri.Host.Split('.');
                if (parts.Length > 0 && parts[0].EndsWith("2"))
                    parts[0] = parts[0].Substring(0, parts[0].Length - 1);
                normalizedSite = $"{uri.Scheme}://{string.Join(".", parts)}{uri.AbsolutePath.TrimEnd('/')}";
            }
            catch { }

            // ── Build PowerShell script using a raw string to avoid escaping hell ─
            // All PS variable signs ($) and braces ({}) are literal PS syntax.
            // We interpolate only the C# values we need.
            string pageNamePs   = EscapePs1(pageNameNoExt);
            string pageTitlePs  = EscapePs1(pageNameNoExt);
            string subfolderPs  = EscapePs1(subfolder);

            var script = new System.Text.StringBuilder();

            // ── Section 1: Header + config ────────────────────────────────────────
            script.AppendLine($@"
# ──────────────────────────────────────────────────────────────────────────────
# You need to save this script in UTF16-BE BOM encoding. This is best done in Notepad++.
# To Install CSOM driver:
# dotnet add package Microsoft.SharePointOnline.CSOM --version 16.1.21812.12000
# ──────────────────────────────────────────────────────────────────────────────
# All functions are declared BEFORE Add-Type so their parameter type
# annotations resolve at call-time from the DLLs loaded below.
# ──────────────────────────────────────────────────────────────────────────────

# ──────────────────────────────────────────────────────────────────────────────
# Creates (or skips) a blank publishing page inside $Folder.
# Identical to the working original — no changes here.
# ──────────────────────────────────────────────────────────────────────────────
function Create-WPPageX($Ctx, $Folder, $pageName, $PageTitle, $pageListUrl, $pageContent) {{
    $sucess = $false
    if (!$pageName.toLower().Contains('.aspx')) {{
        $pageName += '.aspx'
    }}

    $Web = $Ctx.Web
    $ctx.Load($Web)
    $Ctx.ExecuteQuery()

   
    $PageList = $Web.GetList($pageListUrl)
    $Ctx.Load($PageList)
    $Ctx.ExecuteQuery()

    $query = New-Object Microsoft.SharePoint.Client.CamlQuery
    $query.ViewXml = ""
<View>
	<Query>
		<Where>
			<Contains>
				<FieldRef Name='FileLeafRef' />
				<Value Type='Text'>"" +
                     $pageName + 
				""</Value>
			</Contains>
		</Where>
	</Query>
</View>""
	$query.FolderServerRelativeUrl = $pageListUrl
    $listItems = $PageList.GetItems($query)
    $ctx.load($listItems)
    $ctx.executeQuery()

    if ($listItems.Count -eq 0) {{
        write-host ""Create-WPPage: $pageName on $siteURL"" -foregroundcolor Cyan

        # Get PublishingWeb — pass already-loaded $Web, not $Ctx.Web
		#Write-Host ""Ctx Type: $($Ctx.GetType().Assembly.Location)""
		#Write-Host ""Web Type: $($Web.GetType().Assembly.Location)""
		#read-host
        $PublishingWeb = [Microsoft.SharePoint.Client.Publishing.PublishingWeb]::GetPublishingWeb($Ctx, $Web)
        $RootWeb = $Ctx.Site.RootWeb

        $ctx.Load($PublishingWeb)
        $ctx.Load($RootWeb)
        $Ctx.ExecuteQuery()

        # Locate master page gallery by URL — avoids localization issues
        $PageLayoutName        = ""BlankWebPartPage.aspx""
        $masterPageGalleryUrl  = $RootWeb.ServerRelativeUrl.TrimEnd('/') + ""/_catalogs/masterpage""
        #Write-Host ""Master page gallery URL: $masterPageGalleryUrl""

        $MasterPageList = $RootWeb.GetList($masterPageGalleryUrl)
        $CAMLQuery = New-Object Microsoft.SharePoint.Client.CamlQuery
        $CAMLQuery.ViewXml = ""<View><Query><Where><Eq>"" +
                             ""<FieldRef Name='FileLeafRef' />"" +
                             ""<Value Type='Text'>$PageLayoutName</Value>"" +
                             ""</Eq></Where></Query></View>""
        $PageLayouts = $MasterPageList.GetItems($CAMLQuery)
        $Ctx.Load($PageLayouts)
        $Ctx.ExecuteQuery()

        $PageLayoutItem = $PageLayouts[0]
        #write-Host ""PageLayouts Count: $($PageLayouts.count)""
        $Ctx.Load($PageLayoutItem)
        $Ctx.ExecuteQuery()

        #Write-host -f Yellow ""Creating New Page...""
        $PageInfo                    = New-Object Microsoft.SharePoint.Client.Publishing.PublishingPageInformation
        $PageInfo.Name               = $PageName
        if ($Folder){{
			# Unwrap array if Get-SPOFolders returned Object[]
			if ($Folder -is [System.Array])
			{{ 
				$f = $Folder[0] 
			}} 
			else 
			{{ 
				$f = $Folder 
			}}
			$PageInfo.Folder = $f		
		}}
        $PageInfo.PageLayoutListItem = $PageLayoutItem

        $Page = $PublishingWeb.AddPublishingPage($PageInfo)
        $Ctx.ExecuteQuery()

        #Write-host -f Yellow ""Updating Page Content..."" -NoNewline
        $ListItem = $Page.ListItem
        $Ctx.Load($ListItem)
        $Ctx.ExecuteQuery()

        $ListItem[""Title""]                    = $PageTitle
        $ListItem[""PublishingPageContent""]    = $pageContent
        $ListItem.Update()
        $Ctx.ExecuteQuery()

        Write-host -f Yellow ""Checking-In and Publishing the Page..."" -NoNewline
        $ListItem.File.CheckIn([string]::Empty, [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
        $ListItem.File.Publish([string]::Empty)
        $Ctx.ExecuteQuery()
        #Write-host -f Green ""Done!""
        $sucess = $true

    }} else {{
        $srvRelUrl = [String]::Format(""{{0}}/{{1}}/{{2}}"", $PageList.RootFolder.ServerRelativeUrl, $Folder.Name, $PageName)
        write-host ""[Create-WPPageX]: $srvRelUrl Already exists"" -foregroundcolor Yellow
    }}

    return $sucess
}}


# ──────────────────────────────────────────────────────────────────────────────
# Walks the folder tree starting from $Folder.
#
#   No $TargetPath  → diagnostic mode: prints every URL recursively.
#   With $TargetPath → search mode: returns the Folder object whose
#     ServerRelativeUrl ends with the given relative path,
#     e.g. ""Dean/Stage"" or ""Dean/Stage/SubFolder"".
#
# The $Ctx passed here must be the SAME context used to load $Folder,
# so all CSOM types belong to the same assembly instance.
# ──────────────────────────────────────────────────────────────────────────────
Function Get-SPOFolders {{
    param(
        $Ctx,
        $Folder,
        [string] $TargetPath = """"
    )
	if ([string]::IsNullOrEmpty($TargetPath)){{
		#write 118
		return $Folder #[0]
	}}
    Try {{
        $Ctx.Load($Folder.Folders)
        $Ctx.ExecuteQuery()

        foreach ($SubFolder in $Folder.Folders) {{

            if ([string]::IsNullOrEmpty($TargetPath)) {{
                # ── diagnostic mode ──
                #Write-Host $SubFolder.ServerRelativeUrl
                Get-SPOFolders -Ctx $Ctx -Folder $SubFolder

            }} else {{
                # ── search mode ──
                $normalizedTarget  = $TargetPath.TrimStart('/')
				$ServerRelativeUrl = $SubFolder.ServerRelativeUrl
				#Write-Host ""normalizedTarget : $normalizedTarget""
				#Write-Host ""ServerRelativeUrl : $ServerRelativeUrl""
				#read-host
                if ($ServerRelativeUrl -like ""*/$normalizedTarget"") {{
                    return $SubFolder   # exact match — done
                }}

                # recurse deeper; propagate any non-null result upward
                $found = Get-SPOFolders -Ctx $Ctx -Folder $SubFolder -TargetPath $TargetPath
                if ($null -ne $found) {{ return $found }}
            }}
        }}
    }}
    Catch {{
        Write-Host -ForegroundColor Red ""Error walking folder tree: $($_.Exception.Message)""
    }}
    return $null
}}


Function Invoke-LoadMethod() {{
    param(
        [Microsoft.SharePoint.Client.ClientObject] $Object,
        [string] $PropertyName
    )
    $ctx         = $Object.Context
    $load        = [Microsoft.SharePoint.Client.ClientContext].GetMethod(""Load"")
    $type        = $Object.GetType()
    $clientLoad  = $load.MakeGenericMethod($type)

    $Parameter      = [System.Linq.Expressions.Expression]::Parameter(($type), $type.Name)
    $Expression     = [System.Linq.Expressions.Expression]::Lambda(
        [System.Linq.Expressions.Expression]::Convert(
            [System.Linq.Expressions.Expression]::PropertyOrField($Parameter, $PropertyName),
            [System.Object]
        ), $($Parameter))

    $ExpressionArray = [System.Array]::CreateInstance($Expression.GetType(), 1)
    $ExpressionArray.SetValue($Expression, 0)
    $clientLoad.Invoke($ctx, @($Object, $ExpressionArray))
}}


Function Grant-SPOFolderPermission() {{
    Param(
        [Microsoft.SharePoint.Client.Folder] $Folder,
        [String] $UserAccount,
        [String] $PermissionLevel
    )
    Try {{
        $ctx.Load($Folder.ListItemAllFields)
        $Ctx.ExecuteQuery()

        Invoke-LoadMethod -Object $Folder.ListItemAllFields -PropertyName ""HasUniqueRoleAssignments""
        $Ctx.ExecuteQuery()

        If ($Folder.ListItemAllFields.HasUniqueRoleAssignments -ne $true) {{
            $Folder.ListItemAllFields.BreakRoleInheritance($False, $False)
            $Ctx.ExecuteQuery()

            $Folder.ListItemAllFields.RoleAssignments.GetByPrincipal($spCurrentUser).DeleteObject()
            $Folder.Update()
            $Ctx.ExecuteQuery()
            Write-host -f Yellow ""`tFolder's Permission inheritance broken...""
        }}

        $User = $Ctx.Web.EnsureUser($UserAccount)
        $Ctx.load($User)
        $Ctx.ExecuteQuery()

        $Role  = $Ctx.web.RoleDefinitions.GetByName($PermissionLevel)
        $RoleDB = New-Object Microsoft.SharePoint.Client.RoleDefinitionBindingCollection($Ctx)
        $RoleDB.Add($Role)

        $UserPermissions = $Folder.ListItemAllFields.RoleAssignments.Add($User, $RoleDB)
        $Folder.Update()
        $ctx.ExecuteQuery()

        Write-host -f Green ""`tAdded $($User.LoginName) to $($Folder.Name) Permissions!""
    }}
    catch {{
        write-host ""Error in Grant Permissions: $($_.Exception.Message)"" -foregroundcolor Red
    }}
}}


function Remove-Page() {{
    param(
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ClientContext] $Ctx,
        [Parameter(Mandatory=$true)] [string] $FileRelativeURL
    )
    Try {{
        write-host 507 $FileRelativeURL
        $File = $Ctx.Web.GetFileByServerRelativeUrl($FileRelativeURL)
        $Ctx.Load($File)
        $Ctx.ExecuteQuery()
        $File.Name
        $File.DeleteObject()
        $Ctx.ExecuteQuery()
        write-host -f Green ""File has been deleted successfully!""
    }}
    Catch {{
        write-host -f Red ""Error deleting file !"" $_.Exception.Message
    }}
}}
function Test-PageExistsInFolder($targetFolder, $fileName) {{
    
    try {{
		if (!$fileName.contains("".aspx"")){{
			$fileName += "".aspx""
		}}
         if (-not [string]::IsNullOrEmpty($targetFolder)) {{
	 
			$files = $targetFolder.Files
			$Ctx.Load($files)
			$Ctx.ExecuteQuery()
			$fileExists = $false
			forEach($file in $files){{
				#write-host $file.Name
				if ($file.Name -eq $fileName){{
					$fileExists = $true
					break
				}}
			}}
			return $fileExists
		 }}
		 else
		 {{
			return $false 
		 }}
    }}
    catch {{
        Write-Host ""Error checking file: $($_.Exception.Message)"" -ForegroundColor Red
        return $false
    }}
}}
function Get-WebPartsSnapshot
{{
$JsonRaw = @'
{json}
'@


# ── Parse JSON to PowerShell object ─────────────────────────────
$PageSnapshot = $JsonRaw | ConvertFrom-Json
return $PageSnapshot	
}}
function Convert-WebPartPropValue($value, [string[]]$UnixFields = @()) {{
    # Boolean
    $boolResult = $false
    if ([bool]::TryParse($value, [ref]$boolResult)) {{
        return $boolResult
    }}

    # DateTime (standard formats: ISO, RFC, etc.)
    $dateResult = [datetime]::MinValue
    if ([datetime]::TryParse($value, [ref]$dateResult)) {{
        return $dateResult
    }}

    # Integer
    $intResult = 0
    if ([int]::TryParse($value, [ref]$intResult)) {{

        # Unix timestamp heuristic: seconds since epoch, roughly 2001–2100
        $unixMin = 978307200   # 2001-01-01
        $unixMax = 4102444800  # 2100-01-01
        if ($intResult -ge $unixMin -and $intResult -le $unixMax) {{
            return [System.DateTimeOffset]::FromUnixTimeSeconds($intResult).UtcDateTime
        }}

        return $intResult
    }}

    # Fallback — plain string
    return $value
}}
function Update-WebPartsProps
{{
	 param(
	 [Microsoft.SharePoint.Client.WebParts.WebPartDefinition]$WebPart, 
	 [PSCustomObject]$Properties
	) 

try {{
        $ctx = $WebPart.Context

        $wp = $WebPart.WebPart
        $ctx.Load($wp)
        $ctx.Load($wp.Properties)
        $ctx.ExecuteQuery()

        Write-Host ""`n[Update-WebPartsProps] Applying properties..."" -ForegroundColor Yellow

        foreach ($prop in $Properties.PSObject.Properties) {{
            $key       = $prop.Name
            $converted = Convert-WebPartPropValue $prop.Value

            try {{
                $wp.Properties[$key] = $converted
                #Write-Host ""  [$key] = $converted ($($converted.GetType().Name))"" -ForegroundColor Gray
            }} catch {{
                Write-Host ""  WARN: skipping [$key] — $($_.Exception.Message)"" -ForegroundColor Yellow
            }}
        }}

        $WebPart.SaveWebPartChanges()
        $ctx.ExecuteQuery()

        Write-Host ""  Web part properties saved."" -ForegroundColor Green
    }}
    catch {{
        Write-Host ""ERROR in Update-WebPartsProps: $($_.Exception.Message)"" -ForegroundColor Red
        Write-Host ""  Line: $($_.InvocationInfo.ScriptLineNumber)"" -ForegroundColor Yellow
    }}
}}

function Add-WebPartToPage {{
    param(
        [Microsoft.SharePoint.Client.ClientContext] $Ctx,
        $Folder,
        [string] $PageName,
        [string] $WebPartName,
        [string] $WpGalleryRelUrl = ""/home/_catalogs/wp/"",
		$WebPartProps
    )

    $pageRelUrl = $null   # declared here so the catch block can attempt rollback

    try {{
        # ── Resolve URLs ──────────────────────────────────────────────────────
        $Web = $Ctx.Web
        $Ctx.Load($Web)
        $Ctx.ExecuteQuery()

        $hostRoot   = ""https://"" + ([URI]$Ctx.Url).Host
        $pageRelUrl = $Folder.ServerRelativeUrl.TrimEnd('/') + ""/"" + $PageName + "".aspx""

        #Write-Host ""Page : $pageRelUrl""  -ForegroundColor Cyan
        #Write-Host ""WP   : $WebPartName"" -ForegroundColor Cyan

        # ── Step 1: find WebPart in gallery ───────────────────────────────────
        Write-Host ""`n[Step 1] Searching WebPart gallery..."" -ForegroundColor Yellow

        $RootWeb = $Ctx.Site.RootWeb
        $Ctx.Load($RootWeb)
        $Ctx.ExecuteQuery()

        $galleryFolder = $RootWeb.GetFolderByServerRelativeUrl($WpGalleryRelUrl)
        $Ctx.Load($galleryFolder.Properties)
        $Ctx.ExecuteQuery()

        $galleryListId = [System.Guid]::New($galleryFolder.Properties[""vti_listname""].ToString())
        $galleryList   = $RootWeb.Lists.GetById($galleryListId)
        $Ctx.Load($galleryList)
        $Ctx.ExecuteQuery()

        $caml = New-Object Microsoft.SharePoint.Client.CamlQuery
        $caml.ViewXml = ""<View><Query></Query></View>""
        $galleryItems = $galleryList.GetItems($caml)
        $Ctx.Load($galleryItems)
        $Ctx.ExecuteQuery()

        $wpItem = $null
        foreach ($item in $galleryItems) {{
            if ($item[""Title""] -eq $WebPartName) {{ $wpItem = $item; break }}
        }}
        if ($null -eq $wpItem) {{
            Write-Host ""WARNING: WebPart '$WebPartName' not found in gallery — writing error placeholder."" -ForegroundColor Yellow

            # Check out the page so we can write to it
            $pageFile = $Ctx.Web.GetFileByServerRelativeUrl($pageRelUrl)
            $Ctx.Load($pageFile)
            $Ctx.ExecuteQuery()

            try {{
                $pageFile.CheckOut()
                $Ctx.ExecuteQuery()
            }} catch {{
                if ($_.Exception.Message -notmatch ""already checked out"") {{
                    Write-Host ""CheckOut error: $($_.Exception.Message)"" -ForegroundColor Red
                    return $false
                }}
            }}

            # Build error div visible on the page
            $errorDiv = @""
<div style=""border: 2px solid #cc0000; padding: 10px; margin: 6px 0; background: #fff0f0; color: #cc0000; font-weight: bold; font-family: sans-serif;"">
  &#x26A0;&nbsp;Web part not found in gallery: $WebPartName
</div><p><br/></p>
""@

            $Ctx.Load($pageFile.ListItemAllFields)
            $Ctx.ExecuteQuery()
            $pageFields = $pageFile.ListItemAllFields
            $existing   = $pageFields[""PublishingPageContent""]
            $pageFields[""PublishingPageContent""] = if ([string]::IsNullOrEmpty($existing)) {{
                ""<div>$errorDiv</div>""
            }} else {{
                $existing + $errorDiv
            }}
            $pageFields.Update()
            $Ctx.ExecuteQuery()

            # Check in and publish so the page is not left locked
            $pageFile.CheckIn(""Web part '$WebPartName' not found — error placeholder added"",
                [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
            $pageFile.Publish(""Error placeholder for missing web part '$WebPartName'"")
            $Ctx.ExecuteQuery()
            Write-Host ""Error placeholder written and page published."" -ForegroundColor Yellow

            return $false
        }}
        #Write-Host ""Found: '$WebPartName' (ID=$($wpItem.ID))"" -ForegroundColor Green

        $Ctx.Load($wpItem.File)
        $Ctx.ExecuteQuery()

        $wpFileUrl = $hostRoot + $wpItem.File.ServerRelativeUrl
        #Write-Host ""Downloading: $wpFileUrl""

        $response = Invoke-WebRequest $wpFileUrl -UseDefaultCredentials -UseBasicParsing
        if ($response.StatusCode -ne 200) {{
            Write-Host ""ERROR: HTTP $($response.StatusCode)"" -ForegroundColor Red
            return $false
        }}

        $bom   = [char]239 + [char]187 + [char]191
        $raw   = """"
        foreach ($b in $response.Content) {{ $raw += [char]$b }}
        $wpXml = $raw -replace [regex]::Escape($bom), """"

        try   {{ $null = [xml]$wpXml; Write-Host ""XML valid. Length: $($wpXml.Length)"" -ForegroundColor Green }}
        catch {{ Write-Host ""ERROR: invalid XML:`n$wpXml"" -ForegroundColor Red; return $false }}

        # ── Step 2: check out ─────────────────────────────────────────────────
        Write-Host ""`n[Step 2] Checking out page..."" -ForegroundColor Yellow

        $pageFile = $Ctx.Web.GetFileByServerRelativeUrl($pageRelUrl)
        $Ctx.Load($pageFile)
        $Ctx.ExecuteQuery()

        try {{
            $pageFile.CheckOut()
            $Ctx.ExecuteQuery()
            Write-Host ""Checked out."" -ForegroundColor Green
        }} catch {{
            if ($_.Exception.Message -match ""already checked out"") {{
                Write-Host ""Already checked out — continuing."" -ForegroundColor Yellow
            }} else {{
                Write-Host ""CheckOut error: $($_.Exception.Message)"" -ForegroundColor Red
                return $false
            }}
        }}

        # ── Step 3: register in wpz zone ──────────────────────────────────────
        Write-Host ""`n[Step 3] Registering WebPart in 'wpz'..."" -ForegroundColor Yellow

        $wpManager    = $pageFile.GetLimitedWebPartManager(
                            [System.Web.UI.WebControls.WebParts.PersonalizationScope]::Shared)
        $importedDef  = $wpManager.ImportWebPart($wpXml)
        $wpDefinition = $wpManager.AddWebPart($importedDef.WebPart, ""wpz"", 0)
        $Ctx.Load($wpDefinition)
        $Ctx.ExecuteQuery()
		
		if ($WebPartProps){{
			Update-WebPartsProps -WebPart $wpDefinition -Properties $WebPartProps
		}}

        $storageKey = $wpDefinition.Id.ToString(""D"")
        #Write-Host ""WebPart registered. StorageKey = $storageKey"" -ForegroundColor Green

        # ── Step 4: write ms-rte-wpbox placeholder ────────────────────────────
        Write-Host ""`n[Step 4] Writing placeholder into PublishingPageContent..."" -ForegroundColor Yellow

        $placeholder = @""
<div class=""ms-rtestate-read ms-rte-wpbox"" contenteditable=""false"" unselectable=""on"">
  <div class=""ms-rtestate-notify ms-rtestate-read $storageKey"" id=""div_$storageKey"" unselectable=""on"">
  </div>
  <div id=""vid_$storageKey"" unselectable=""on"" style=""display: none;"">
  </div>
</div>
""@

        $Ctx.Load($pageFile.ListItemAllFields)
        $Ctx.ExecuteQuery()
        $pageFields = $pageFile.ListItemAllFields

        $existing   = $pageFields[""PublishingPageContent""]
        $newContent = if ([string]::IsNullOrEmpty($existing)) {{
            ""<div>"" + $placeholder + ""</div>""
        }} else {{
            $existing + $placeholder + ""<p><br/></p>""
        }}

        $pageFields[""PublishingPageContent""] = $newContent
        $pageFields.Update()
        $Ctx.ExecuteQuery()
        #Write-Host ""PublishingPageContent written."" -ForegroundColor Green

        # ── Step 5: check in and publish ──────────────────────────────────────
        Write-Host ""`n[Step 5] Publishing..."" -ForegroundColor Yellow

        $pageFile.CheckIn(""Added WebPart '$WebPartName'"",
            [Microsoft.SharePoint.Client.CheckinType]::MajorCheckIn)
        $pageFile.Publish(""Published after adding '$WebPartName'"")
        $Ctx.ExecuteQuery()
        Write-Host ""Published successfully."" -ForegroundColor Green

        return $true

    }} catch {{
        Write-Host ""ERROR: $($_.Exception.Message)"" -ForegroundColor Red
        Write-Host ""Line : $($_.InvocationInfo.ScriptLineNumber)"" -ForegroundColor Yellow

        # Attempt to roll back checkout so the page isn't left locked
        if ($null -ne $pageRelUrl) {{
            try {{
                $f = $Ctx.Web.GetFileByServerRelativeUrl($pageRelUrl)
                $Ctx.Load($f)
                $Ctx.ExecuteQuery()
                $f.CheckIn(""Rollback"", [Microsoft.SharePoint.Client.CheckinType]::MinorCheckIn)
                $Ctx.ExecuteQuery()
                Write-Host ""Checkout rolled back."" -ForegroundColor Yellow
            }} catch {{}}
        }}

        return $false
    }}
}}
Function Create-FolderX($ctx, $LibName, $pathToCreate)
{{
    # Валидация входных параметров
    if ($null -eq $ctx) {{
        Write-Host ""ERROR: ctx is null"" -ForegroundColor Red
        return $null
    }}
    if ([string]::IsNullOrWhiteSpace($LibName)) {{
        Write-Host ""ERROR: LibName is empty or null"" -ForegroundColor Red
        return $null
    }}

    # Санация: удаление недопустимых символов
    $illegalChars = '[#%&*:<>?{{|}}~]'
    if ($LibName -match $illegalChars) {{ Write-Host ""ERROR: LibName contains illegal characters"" -ForegroundColor Red; return $null }}
    
    # Если путь пустой, возвращаем корневую папку библиотеки
    if ([string]::IsNullOrWhiteSpace($pathToCreate)) {{
        $rootFolder = $ctx.Web.GetFolderByServerRelativeUrl($LibName.Trim())
        $ctx.Load($rootFolder)
        $ctx.ExecuteQuery()
        return $rootFolder
    }}

    if ($pathToCreate -match $illegalChars) {{ 
        Write-Host ""ERROR: pathToCreate contains illegal characters $pathToCreate"" -ForegroundColor Red
        return $null 
    }}

    # Нормализуем слеши и разбиваем путь на массив каталогов
    # Используем как прямой, так и обратный слеш для разделения
    $folderParts = $pathToCreate.Trim().Replace('\', '/').Split('/', [System.StringSplitOptions]::RemoveEmptyEntries)

    # Начинаем от корня библиотеки
    $currentPath = $LibName.Trim()
    
    # Переменная для хранения ссылки на последнюю созданную папку
    $lastFolder = $null

    try {{
        # Идем по каждому элементу пути сверху вниз
        foreach ($part in $folderParts) {{
            # Формируем путь для текущего уровня
            $currentPath = ($currentPath + ""/"" + $part).Replace('//', '/')
            #Write-Host ""Checking/Creating folder level: $currentPath"" -ForegroundColor Yellow

            # Используем встроенный механизм CSOM для добавления папки по относительному пути.
            # Метод .Add() на коллекции Folders веб-сайта умеет создавать элемент, 
            # но для цепочки безопаснее создавать их последовательно в цикле.
            $lastFolder = $ctx.Web.Folders.Add($currentPath)
            $ctx.Load($lastFolder)
            $ctx.ExecuteQuery()
        }}

        #Write-Host ""SUCCESS: Full path created: $($lastFolder.ServerRelativeUrl)"" -ForegroundColor Green
        return $lastFolder
    }}
    catch {{
        Write-Host ""ERROR creating folder structure at '$currentPath': $($_.Exception.Message)"" -ForegroundColor Red
        return $null
    }}
}}
Function Create-Folder($ctx, $LibName, $pathToCreate)
{{
    # Validate input parameters
	write-host $LibName $pathToCreate
    if ($null -eq $ctx) {{
        Write-Host ""ERROR: ctx is null"" -ForegroundColor Red
        return $null
    }}
    if ([string]::IsNullOrWhiteSpace($LibName)) {{
        Write-Host ""ERROR: LibName is empty or null"" -ForegroundColor Red
        return $null
    }}
    #if ([string]::IsNullOrWhiteSpace($pathToCreate)) {{
    #    Write-Host ""ERROR: pathToCreate is empty or null"" -ForegroundColor Red
    #    return $null
    #}}

    # Sanitize: remove illegal SharePoint characters
    $illegalChars = '[#%&*:<>?{{|}}~]'
    if ($LibName     -match $illegalChars) {{ Write-Host ""ERROR: LibName contains illegal characters""     -ForegroundColor Red; return $null }}
    if ($pathToCreate -match $illegalChars) {{ 
		Write-Host ""ERROR: pathToCreate contains illegal characters $pathToCreate"" -ForegroundColor Red; 
		#read-host
		return $null 
		}}

    # Build full path
    $fullPath = (Join-Path $LibName.Trim() $pathToCreate.Trim()).Replace('\', '/')
    Write-Host $fullPath -f Yellow
	
	
 
	#try {{
		$newFolder = $ctx.Web.Folders.Add($fullPath)
		$ctx.ExecuteQuery()
		Write-Host ""Folder created: $($newFolder.ServerRelativeUrl)"" -ForegroundColor Green
		return $newFolder
	#}}
	#catch {{
	#	Write-Host ""ERROR creating folder '$fullPath': $($_.Exception.Message)"" -ForegroundColor Red
	#	return $null
	#}}
    
}}

# ══════════════════════════════════════════════════════════════════════════════
# DLL loading — must come AFTER function declarations so type annotations
# in params resolve against these assemblies at call-time.
# ══════════════════════════════════════════════════════════════════════════════
Add-Type -Path ""C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.dll""
Add-Type -Path ""C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Runtime.dll""
Add-Type -Path ""C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.Portable.dll""
Add-Type -Path ""C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Taxonomy.dll""
Add-Type -Path ""C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Publishing.dll""
Add-Type -Path ""C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.Search.dll""
Add-Type -Path ""C:\Program Files\Common Files\Microsoft Shared\Web Server Extensions\16\ISAPI\Microsoft.SharePoint.Client.UserProfiles.dll""


# ══════════════════════════════════════════════════════════════════════════════
# Entry point
# ══════════════════════════════════════════════════════════════════════════════
#. ""..\Utils-Request.ps1""
#. ""..\Utils-DualLanguage.ps1""

start-transcript ""1.AddPages.log""

$Credentials = Get-Credential

# ── Configuration ──────────────────────────────────────────────────────────────
#$pageName    = ""candidate_add_reviewer""

# Relative paths inside the Pages library — any nesting depth is supported.
# Examples:  ""Dean""  |  ""Dean/Stage""  |  ""Dean/Stage/SubFolder""
#$TargetFolderPaths = @(
#    ""Dean/Stage""
#)
#$TargetFolderPaths = @(
#    ""Dean"","""",""Dean/Stage"",""Reports/V2""
#)


$PageSnapshot = Get-WebPartsSnapshot

# ── Connect ────────────────────────────────────────────────────────────────────
$siteUrl =   $PageSnapshot.Site
$ListName    = ""Pages""
$pageName    = $PageSnapshot.Page.Replace("".aspx"","""")
$TargetFolderPaths = @($PageSnapshot.Subfolder)
$pageContent = ""<h1>Test Page</h1>""

write-host ""URL: $siteUrl"" -foregroundcolor Yellow

[Microsoft.SharePoint.Client.ClientContext]$Ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)
$Ctx.Credentials = $Credentials

$Web = $Ctx.Web
$Ctx.Load($Web)
$Ctx.ExecuteQuery()

# Load Pages root folder once — reused by Get-SPOFolders for every target path
$List = $Ctx.Web.Lists.GetByTitle($ListName)

$urlCSV = @()

foreach ($TargetFolderPath in $TargetFolderPaths) {{
    Write-Host ""Create Folder $TargetFolderPath"" -f Yellow
    $xNull = Create-FolderX $Ctx $ListName $TargetFolderPath
    
	# Load only the properties we actually need — avoids ""StorageMetrics does not exist""
	$Ctx.Load($List)
	$Ctx.ExecuteQuery()
	$Ctx.Load($List.RootFolder)   # load RootFolder together with List
	$Ctx.ExecuteQuery()           # one round-trip — RootFolder.ServerRelativeUrl now populated
    
	# Now get folder by URL — no StorageMetrics issue
	$listRelUrl = $List.RootFolder.ServerRelativeUrl
	$pagesRootFolder = $Ctx.Web.GetFolderByServerRelativeUrl($listRelUrl)
	$Ctx.Load($pagesRootFolder)
	$Ctx.ExecuteQuery()
	 
	Write-Host ""Get target folder: '$TargetFolderPath' ──"" -ForegroundColor Yellow
	#read-host
    # ── Find the folder at any depth using Get-SPOFolders ──
    $targetFolder = Get-SPOFolders -Ctx $Ctx -Folder $pagesRootFolder -TargetPath $TargetFolderPath

    if ($null -eq $targetFolder) {{
        Write-Host ""Folder '$TargetFolderPath' not found in '$ListName' — skipping."" -ForegroundColor Red
        continue
    }}

	Write-Host ""Test Page $pageName exists in folder"" -ForegroundColor Magenta
	$fileExists = Test-PageExistsInFolder $targetFolder $pageName

    # ── Compute URLs once — needed in all branches ───────────────────────────
    # targetFolder.ServerRelativeUrl already contains the full path, e.g.
    #   /home/huca/.../Pages/Dean/Stage
    $pageListUrl = $targetFolder.ServerRelativeUrl   # passed to Web.GetList() and Remove-Page
    $SiteHost    = $([uri]$siteUrl).Host
    $pagePath    = ""$pageListUrl/$pageName.aspx""
    $fileRelUrl  = $pagePath                         # server-relative path for Remove-Page

    # ── Decision phase: set flags only, no execution here ───────────────────
    $shouldCreatePage  = $false
    $shouldAddWebParts = $false

    if ($fileExists) {{
        $fullPageUrl  = ""https://$SiteHost$pagePath""
        Write-Host ""  Page: $fullPageUrl"" -ForegroundColor White
        $deleteChoice = Read-Host ""Page '$pageName' already exists in '$TargetFolderPath'. Delete and recreate? (Y/N)""

        if ($deleteChoice -eq 'Y' -or $deleteChoice -eq 'y') {{
            Write-Host ""Deleting existing page: $fileRelUrl"" -ForegroundColor Yellow
            Remove-Page -Ctx $Ctx -FileRelativeURL $fileRelUrl
            $shouldCreatePage  = $true
            $shouldAddWebParts = $true
        }}
        else {{
            # Warn the user: show full page path and list of web parts to be added
            Write-Host """"
            Write-Host ""WARNING: Page will NOT be deleted."" -ForegroundColor Yellow
            Write-Host ""  Page: $fullPageUrl"" -ForegroundColor White
            Write-Host ""The following web parts will be added to the existing page:"" -ForegroundColor Yellow
            $i = 1
            forEach ($wPart in $PageSnapshot.WebParts) {{
                Write-Host ""  $i. $($wPart.Title)  [Zone: $($wPart.Zone), Position: $($wPart.Position)]"" -ForegroundColor Cyan
                $i++
            }}
            Write-Host """"
            $addWpChoice = Read-Host ""Add web parts to this page? (Y/N)""
            if ($addWpChoice -eq 'Y' -or $addWpChoice -eq 'y') {{
                $shouldAddWebParts = $true
            }}
            else {{
                Write-Host ""Skipping page '$fullPageUrl'."" -ForegroundColor Gray
            }}
            # $shouldCreatePage stays $false — page is kept as-is
        }}
    }}
    else {{
        Write-Host ""Page does not exist. Creating page with WebParts...""
        $shouldCreatePage  = $true
        $shouldAddWebParts = $true
    }}

    # ── Execution phase: each action runs in exactly one place ───────────────
    if ($shouldCreatePage) {{
        $urlItem     = """" | Select URL
        $urlItem.URL = ""https://$SiteHost$pagePath""
        $urlCSV     += $urlItem

        Create-WPPageX $Ctx $targetFolder $pageName $targetFolder.Name $pageListUrl $pageContent
    }}

    if ($shouldAddWebParts) {{
        forEach ($wPart in $PageSnapshot.WebParts) {{
            $wpTitle = $wPart.Title
            Write-Host ""Adding web part '$wpTitle' to page...""
            $ok = Add-WebPartToPage -Ctx $Ctx -PageName $pageName `
                -folder $targetFolder `
                -WebPartName $wpTitle `
                -WebPartProps $wPart.Properties
        }}
    }}

}}

$urlCSV | Export-Csv -Path URLList.csv -Delimiter ""`t"" -NoTypeInformation -Encoding Default
stop-transcript

");

            // ── Show in UniversalPreviewWindow ────────────────────────────────
            var win = new SPUtil.App.Views.UniversalPreviewWindow
            {
                Title  = $"PowerShell Export — {pageName}",
                Owner  = Application.Current.MainWindow,
                Width  = 1100,
                Height = 750
            };

            string scriptText = script.ToString();

            var buttons = new System.Collections.ObjectModel.ObservableCollection<DialogButton>
            {
                new DialogButton
                {
                    Caption = "📋  Copy script",
                    Action  = () =>
                    {
                        try
                        {
                            System.Windows.Clipboard.SetText(scriptText);
                            StatusMessage = $"✔ PowerShell script copied ({wpList.Count} WebParts)";
                        }
                        catch (Exception ex)
                        {
                            _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                            StatusMessage = $"Clipboard error: {ex.Message}";
                        }
                    }
                },
                new DialogButton
                {
                    Caption  = "Close",
                    IsCancel = true,
                    Action   = () => win?.Close()
                }
            };

            var vm = new PowerShellPreviewViewModel(scriptText, pageName, wpList.Count, win, buttons);
            win.DataContext = vm;
            win.ShowDialog();

            StatusMessage = $"PowerShell script generated for {pageName} ({wpList.Count} WebParts)";
        }


        /// <summary>Escapes a string value for PowerShell single-quoted string.</summary>
        private static string EscapePs1(string value) =>
            string.IsNullOrEmpty(value) ? "" : value.Replace("'", "''");

        /// <summary>Escapes a string value for JSON embedding.</summary>
      private static string EscapeJson(string value)
        {
            if (string.IsNullOrEmpty(value)) return "";
            return value
                .Replace("\\", "\\\\")
                .Replace("\"", "\\\"")
                .Replace("\r\n", "\\n")
                .Replace("\n", "\\n")
                .Replace("\r", "\\n")
                .Replace("\t", "\\t");
        }

        // ── Data loading ──────────────────────────────────────────────────────
        public async Task LoadDataAsync(string siteUrl, string listId)
        {
            _siteUrl = siteUrl;
            try
            {
                StatusMessage = "Loading pages...";
                var data = await _spService.GetPageItemsAsync(siteUrl, listId);

                // Show only pages (.aspx files), skip folders, sort by Name
                var pages = data
                    .Where(f => !f.IsFolder)
                    .OrderBy(f => f.Name, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                Pages = new ObservableCollection<SPFileData>(pages);
                StatusMessage = $"Pages: {pages.Count}";
            }
            catch (Exception ex)
            {
                _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                StatusMessage = $"Load error: {ex.Message}";
            }
        }

        private async Task LoadWebPartsAsync(string fileUrl)
        {
            _log.Debug("LoadWebParts: {Url}", fileUrl);
            try
            {
                StatusMessage = "Load Web parts...";
                WebParts.Clear();
                // GetWebPartsWithPositionAsync — enriches result with VisualPosition
                // without modifying SharePointService.cs
                var wpData = await _spService.GetWebPartsWithPositionAsync(_siteUrl, fileUrl);
                WebParts = new ObservableCollection<SPWebPartData>(wpData);
                ExportWpToPowerShellCommand.RaiseCanExecuteChanged();
                _log.Information("WebParts loaded: {Count} for {Url}", wpData.Count, fileUrl);
                StatusMessage = WebParts.Any()
                    ? $"Web parts count: {WebParts.Count}"
                    : "No Web parts found.";
            }
            catch (Exception ex)
            {
                _log.Error(ex, "LoadWebParts failed for {Url}", fileUrl);
                StatusMessage = $"Web part error: {ex.Message}";
                Debug.WriteLine(ex.ToString());
            }
        }
    }
}
