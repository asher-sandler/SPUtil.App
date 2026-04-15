using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Infrastructure;
using SPUtil.Services;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;

namespace SPUtil.App.ViewModels
{
    public class Library101ViewModel : BindableBase
    {
        private readonly ISharePointService _spService;
        private ObservableCollection<SPFileData> _files = new();
        private string _statusMessage = string.Empty;
        private string _libraryTitle = string.Empty;

        public string LibraryTitle { get => _libraryTitle; set => SetProperty(ref _libraryTitle, value); }
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public ObservableCollection<SPFileData> Files { get => _files; set => SetProperty(ref _files, value); }
		// признак что это source
		private bool _isSourceMode;
		public bool IsSourceMode 
		{ 
			get => _isSourceMode; 
			set => SetProperty(ref _isSourceMode, value); 
		}

        // Команды
        public DelegateCommand SelectAllCommand { get; }
        public DelegateCommand CopySelectedCommand { get; }
        public DelegateCommand DeleteSelectedCommand { get; }

        public Library101ViewModel(ISharePointService spService)
        {
            _spService = spService;

            // Инициализация команды выделения всех элементов
            SelectAllCommand = new DelegateCommand(() =>
            {
                if (Files == null) return;
                foreach (var f in Files) f.IsSelected = true;
                
                // Переприсваиваем коллекцию для уведомления UI об изменениях внутри списка
                var temp = new ObservableCollection<SPFileData>(Files);
                Files = temp;
            });

            // Инициализация команды копирования
            CopySelectedCommand = new DelegateCommand(() => {
                var selectedCount = Files?.Count(f => f.IsSelected) ?? 0;
                StatusMessage = $"Копирование {selectedCount} элементов...";
            });

            // Инициализация команды удаления
            DeleteSelectedCommand = new DelegateCommand(() => {
                var selectedCount = Files?.Count(f => f.IsSelected) ?? 0;
                StatusMessage = $"Удаление {selectedCount} элементов...";
            });
        }

        public async Task LoadDataAsync(string siteUrl, string listId)
        {
            try
            {
                StatusMessage = "Загрузка данных из SharePoint...";
                string cleanId = listId.StartsWith("id:") ? listId.Substring(3) : listId;
                
                var data = await _spService.GetLibraryItemsAsync(siteUrl, cleanId);

                if (data.Count > 250)
                {
                    StatusMessage = $"Внимание: в библиотеке {data.Count} элементов. Показаны только первые 250.";
                    Files = new ObservableCollection<SPFileData>(data.Take(250));
                }
                else
                {
                    StatusMessage = $"Всего элементов: {data.Count}";
                    Files = new ObservableCollection<SPFileData>(data);
                }
            }
            catch (Exception ex) 
            {
                StatusMessage = $"SERVER ERROR: {ex.Message}"; 
                System.Diagnostics.Debug.WriteLine($"Полная ошибка: {ex.ToString()}");
            }
        }
    }
}