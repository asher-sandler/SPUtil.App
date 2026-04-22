using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Infrastructure;
using SPUtil.Services;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;

namespace SPUtil.App.ViewModels
{
    public class PagesViewModel : BindableBase
    {
        private readonly ISharePointService _spService;
        private string  _siteUrl       = string.Empty;
        private string  _targetSiteUrl = string.Empty;
        private string  _statusMessage = "Готов";
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
                StatusMessage = "Данные свойств выведены в Output";
            });

            ShowWebPartsPreviewCommand = new DelegateCommand(() =>
            {
                if (SelectedPage == null)
                {
                    MessageBox.Show("Выберите страницу в списке сверху.",
                        "Нет выбранной страницы", MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }
                if (!WebParts.Any())
                {
                    MessageBox.Show("На выбранной странице нет веб-частей или они ещё не загружены.",
                        "Нет веб-частей", MessageBoxButton.OK, MessageBoxImage.Information);
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
        }

        // ── Called by MainWindowViewModel after creating this VM ──────────────
        public void SetTargetSiteUrl(string url) => _targetSiteUrl = url;


        // ═══════════════════════════════════════════════════════════════════════
        //  Copy Page
        // ═══════════════════════════════════════════════════════════════════════
        private async Task ExecuteCopyPageAsync()
        {
            if (SelectedPage == null)
            {
                MessageBox.Show("Выберите страницу для копирования.",
                    "Нет выбранной страницы", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrEmpty(_targetSiteUrl))
            {
                MessageBox.Show("Подключитесь к целевому сайту (правая панель).",
                    "Нет целевого сайта", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                snapshot = await _spService.GetPageSnapshotAsync(_siteUrl, SelectedPage.FullPath);
            }
            catch (Exception ex)
            {
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
                await _spService.CreatePageFromSnapshotAsync(
                    _targetSiteUrl, targetName, snapshot, subfolderPath);
            }
            catch (Exception ex)
            {
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
            if (SelectedPage == null)
            {
                MessageBox.Show("Выберите страницу для удаления.",
                    "Нет выбранной страницы", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                await _spService.DeletePageAsync(_siteUrl, SelectedPage.Name);

                var removed = Pages.FirstOrDefault(p => p.FullPath == SelectedPage.FullPath);
                if (removed != null) Pages.Remove(removed);

                SelectedPage  = null;
                StatusMessage = "Page deleted.";
            }
            catch (Exception ex)
            {
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
            if (SelectedPage == null)
            {
                MessageBox.Show("Выберите страницу для переименования.",
                    "Нет выбранной страницы", MessageBoxButton.OK, MessageBoxImage.Warning);
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
            if (SelectedPage == null)
            {
                MessageBox.Show("Выберите страницу для сравнения.",
                    "Нет выбранной страницы", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if (string.IsNullOrEmpty(_targetSiteUrl))
            {
                MessageBox.Show("Подключитесь к целевому сайту (правая панель).",
                    "Нет целевого сайта", MessageBoxButton.OK, MessageBoxImage.Warning);
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
            if (SelectedPage == null)
            {
                MessageBox.Show("Выберите страницу для синхронизации.",
                    "Нет выбранной страницы", MessageBoxButton.OK, MessageBoxImage.Warning);
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
                StatusMessage = $"Load error: {ex.Message}";
            }
        }

        private async Task LoadWebPartsAsync(string fileUrl)
        {
            try
            {
                StatusMessage = "Загрузка веб-частей...";
                WebParts.Clear();
                // GetWebPartsWithPositionAsync — enriches result with VisualPosition
                // without modifying SharePointService.cs
                var wpData = await _spService.GetWebPartsWithPositionAsync(_siteUrl, fileUrl);
                WebParts = new ObservableCollection<SPWebPartData>(wpData);
                StatusMessage = WebParts.Any()
                    ? $"Web parts count: {WebParts.Count}"
                    : "No Web parts found.";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка веб-частей: {ex.Message}";
                Debug.WriteLine(ex.ToString());
            }
        }
    }
}
