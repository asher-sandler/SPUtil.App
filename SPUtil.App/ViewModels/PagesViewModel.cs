using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Infrastructure;
using SPUtil.Services;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using System.Diagnostics;

namespace SPUtil.App.ViewModels
{
    public class PagesViewModel : BindableBase
    {
        private readonly ISharePointService _spService;
        private string _siteUrl = string.Empty;
        private string _statusMessage = "Готов";
        private ObservableCollection<SPFileData> _pages = new();
        private ObservableCollection<SPWebPartData> _webParts = new();
        private SPFileData? _selectedPage;

        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public ObservableCollection<SPFileData> Pages { get => _pages; set => SetProperty(ref _pages, value); }
        public ObservableCollection<SPWebPartData> WebParts { get => _webParts; set => SetProperty(ref _webParts, value); }
		// признак что это source
		private bool _isSourceMode;
		public bool IsSourceMode 
		{ 
			get => _isSourceMode; 
			set => SetProperty(ref _isSourceMode, value); 
		}

        // Свойство для отслеживания выбора страницы в верхнем списке
        public SPFileData? SelectedPage
        {
            get => _selectedPage;
            set
            {
                if (SetProperty(ref _selectedPage, value) && value != null)
                {
                    if (!value.IsFolder)
                    {
                        _ = LoadWebPartsAsync(value.FullPath);
                    }
                    else
                    {
                        WebParts.Clear();
                        StatusMessage = "Выбрана папка";
                    }
                }
            }
        }

        public DelegateCommand GetAllPropertiesCommand { get; }

        public PagesViewModel(ISharePointService spService)
        {
            _spService = spService;

            GetAllPropertiesCommand = new DelegateCommand(() =>
            {
                Debug.WriteLine($">>> [Pages] Запрос всех свойств для {WebParts.Count} веб-партов");
                foreach (var wp in WebParts)
                {
                    Debug.WriteLine($"Веб-парт: {wp.Title} ({wp.Id})");
                    foreach (var prop in wp.Properties)
                    {
                        Debug.WriteLine($"   {prop.Key}: {prop.Value}");
                    }
                }
                StatusMessage = "Данные свойств выведены в Output";
            });
        }

        public async Task LoadDataAsync(string siteUrl, string listId)
        {
            _siteUrl = siteUrl;
            try
            {
                StatusMessage = "Загрузка страниц (рекурсивно)...";
                // Метод GetPageItemsAsync должен поддерживать рекурсию в Service
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
                StatusMessage = WebParts.Any() ? $"Найдено веб-частей: {WebParts.Count}" : "Веб-частей не найдено";
            }
            catch (Exception ex)
            {
                StatusMessage = $"Ошибка веб-частей: {ex.Message}";
                Debug.WriteLine(ex.ToString());
            }
        }
    }
}