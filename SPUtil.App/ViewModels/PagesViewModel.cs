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
        private string _siteUrl      = string.Empty;
        private string _statusMessage = "Готов";
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

        // ── Commands ─────────────────────────────────────────────────────────
        public DelegateCommand GetAllPropertiesCommand    { get; }

        /// <summary>
        /// Opens UniversalPreviewWindow with all WebPart properties for the
        /// currently selected page. The window also has a "Copy all" button
        /// that copies the formatted text to the clipboard.
        /// </summary>
        public DelegateCommand ShowWebPartsPreviewCommand { get; }

        public PagesViewModel(ISharePointService spService)
        {
            _spService = spService;

            // Original command — keeps writing to Output window
            GetAllPropertiesCommand = new DelegateCommand(() =>
            {
                Debug.WriteLine($">>> [Pages] {WebParts.Count} web parts for '{SelectedPage?.Name}'");
                foreach (var wp in WebParts)
                {
                    Debug.WriteLine($"WebPart: {wp.Title} ({wp.Id})  Type: {wp.Type}");
                    foreach (var prop in wp.Properties)
                        Debug.WriteLine($"   {prop.Key}: {prop.Value}");
                }
                StatusMessage = "Данные свойств выведены в Output";
            });

            // New command — opens UniversalPreviewWindow
            ShowWebPartsPreviewCommand = new DelegateCommand(() =>
            {
                if (SelectedPage == null)
                {
                    MessageBox.Show(
                        "Выберите страницу в списке сверху.",
                        "Нет выбранной страницы",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                if (!WebParts.Any())
                {
                    MessageBox.Show(
                        "На выбранной странице нет веб-частей или они ещё не загружены.",
                        "Нет веб-частей",
                        MessageBoxButton.OK,
                        MessageBoxImage.Information);
                    return;
                }

                // Create the preview window and its own ViewModel.
                // We pass the window reference into the VM so the "Close" button
                // can close exactly this instance without searching Application.Windows.
                var win = new SPUtil.App.Views.UniversalPreviewWindow
                {
                    Title  = $"WebParts — {SelectedPage.Name}",
                    Owner  = Application.Current.MainWindow,
                    Width  = 1000,
                    Height = 680
                };

                var vm = new WebPartsPreviewViewModel(
                    webParts:    WebParts,
                    pageTitle:   SelectedPage.Name,
                    ownerWindow: win);

                win.DataContext = vm;
                win.ShowDialog();
            });
        }

        // ── Data loading ─────────────────────────────────────────────────────
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
