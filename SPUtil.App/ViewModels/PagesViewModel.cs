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
        private SPFileData? _selectedPage;

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
                {
                    if (!value.IsFolder)
                        _ = LoadWebPartsAsync(value.FullPath);
                    else
                    {
                        WebParts.Clear();
                        StatusMessage = "Выбрана папка";
                    }
                }
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

            string sourceName = System.IO.Path.GetFileNameWithoutExtension(SelectedPage.Name);
            string sourceInfo = $"Page : {SelectedPage.Name}\nPath : {SelectedPage.FullPath}\nSite : {_siteUrl}";

            // ── Step 1: Name + conflict resolution dialog ─────────────────────
            var dialog = new SPUtil.Views.CopyPageDialog(sourceName, _targetSiteUrl, sourceInfo)
            {
                Owner = Application.Current.MainWindow
            };
            if (dialog.ShowDialog() != true) return;

            string targetName   = dialog.TargetPageName;
            string existsAction = dialog.ExistsAction;   // "Replace" | "Rename"

            // ── Step 2: Check existence ───────────────────────────────────────
            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = Application.Current.MainWindow
            };
            infoWin.Show();
            infoWin.UpdateMessage("Checking target site...");

            bool exists = await _spService.PageExistsAsync(_targetSiteUrl, targetName);

            if (exists)
            {
                if (existsAction == "Rename")
                {
                    string oldName = targetName + "_old";
                    infoWin.UpdateMessage($"Renaming existing '{targetName}' → '{oldName}'...");
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
                else  // Replace
                {
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

            // ── Step 4: Create page on target ─────────────────────────────────
            int wpCount = snapshot.WebParts.Count;
            infoWin.UpdateMessage(
                $"Creating '{targetName}' with {wpCount} WebPart(s)...");
            try
            {
                await _spService.CreatePageFromSnapshotAsync(
                    _targetSiteUrl, targetName, snapshot);
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


        // ── Data loading ──────────────────────────────────────────────────────
        public async Task LoadDataAsync(string siteUrl, string listId)
        {
            _siteUrl = siteUrl;
            try
            {
                StatusMessage = "Загрузка страниц (рекурсивно)...";
                var data = await _spService.GetPageItemsAsync(siteUrl, listId);
                Pages = new ObservableCollection<SPFileData>(data);
                StatusMessage = $"Загружено элементов: {data.Count}";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка загрузки: {ex.Message}";
            }
        }

        private async Task LoadWebPartsAsync(string fileUrl)
        {
            try
            {
                StatusMessage = "Загрузка веб-частей...";
                WebParts.Clear();
                var wpData = await _spService.GetWebPartsAsync(_siteUrl, fileUrl);
                WebParts = new ObservableCollection<SPWebPartData>(wpData);
                StatusMessage = WebParts.Any()
                    ? $"Найдено веб-частей: {WebParts.Count}"
                    : "Веб-частей не найдено";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка веб-частей: {ex.Message}";
                Debug.WriteLine(ex.ToString());
            }
        }
    }
}
