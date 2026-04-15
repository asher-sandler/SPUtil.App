using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Services;
using SPUtil.Infrastructure; 
using System.Collections.ObjectModel;
using System.Collections.Generic;
using System.Threading.Tasks;
using System.Linq;
using System.Diagnostics; // Добавлено для Debug.WriteLine
using System;

namespace SPUtil.App.ViewModels
{
    public class List100ViewModel : BindableBase
    {
        private readonly ISharePointService _spService;
        private string _listTitle = string.Empty;
        private string _statusMessage = "Готов"; // Новое свойство для статуса
        private ObservableCollection<SPViewData> _views = new();
        private ObservableCollection<SPFieldData> _fields = new();

        public string ListTitle { get => _listTitle; set => SetProperty(ref _listTitle, value); }
        
        // Свойство для отображения статуса в UI
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }

		private ObservableCollection<SPListItemData> _items = new();
		public ObservableCollection<SPListItemData> Items 
		{ 
			get => _items; 
			set => SetProperty(ref _items, value); 
		}		
		// признак что это source
		private bool _isSourceMode;
		public bool IsSourceMode 
		{ 
			get => _isSourceMode; 
			set => SetProperty(ref _isSourceMode, value); 
		}
        public ObservableCollection<SPFieldData> Fields 
        { 
            get => _fields; 
            set => SetProperty(ref _fields, value); 
        }

        public ObservableCollection<SPViewData> Views 
        { 
            get => _views; 
            set => SetProperty(ref _views, value); 
        }

		private SPViewData _selectedView;
		public SPViewData SelectedView
		{
			get => _selectedView;
			set => SetProperty(ref _selectedView, value);
		}
        // КОМАНДЫ
        public DelegateCommand CreateOnTargetCommand { get; }
        public DelegateCommand CopyWithDataCommand { get; }
        public DelegateCommand CopyViewsCommand { get; }
        public DelegateCommand CompareCommand { get; }

        public List100ViewModel(ISharePointService spService)
        {
            _spService = spService;

            // Инициализация команд с логированием
            CreateOnTargetCommand = new DelegateCommand(() => 
            {
                LogAndStatus("Нажата кнопка: Создать структуру на целевом сайте");
            });

            CopyWithDataCommand = new DelegateCommand(() => 
            {
                LogAndStatus("Нажата кнопка: Копировать вместе с данными");
            });

            CopyViewsCommand = new DelegateCommand(() => 
            {
                LogAndStatus("Нажата кнопка: Копировать представления (Views)");
            });

            CompareCommand = new DelegateCommand(() => 
            {
                LogAndStatus("Нажата кнопка: Сравнить списки");
            });
        }

        // Вспомогательный метод для лога и статуса
        private void LogAndStatus(string message)
        {
            StatusMessage = message;
            Debug.WriteLine($">>> [List100] {DateTime.Now:HH:mm:ss} - {message}");
        }

        public async Task LoadDataAsync(string siteUrl, string listPath)
        {
            LogAndStatus($"Загрузка данных для списка: {listPath}...");
            
            Fields.Clear();
            Views.Clear();

            string cleanId = listPath.StartsWith("id:") ? listPath.Substring(3) : listPath;

            try
            {
                var fieldsData = await _spService.GetListFieldsAsync(siteUrl, cleanId);
                var result = fieldsData
                     .Where(f =>
                         // 1. Разрешаем поля на иврите (начинаются с _x) 
                         // ИЛИ разрешаем те, что НЕ начинаются с подчеркивания
                         (f.InternalName.StartsWith("_x") || !f.InternalName.StartsWith("_")) &&

                         // 2. Исключаем технические поля
                         f.TypeAsString != "Computed" &&

                         // 3. Исключаем конкретные системные ID
                         f.InternalName != "ContentTypeId" &&
                         f.InternalName != "Attachments"
                     )
                     .ToList();
                Fields = new ObservableCollection<SPFieldData>(result);
                LogAndStatus($"Загружено полей: {Fields.Count}");
            }
            catch (Exception ex)
            {
                LogAndStatus($"Ошибка загрузки полей: {ex.Message}");
            }

            try
            {
                var viewsData = await _spService.GetListViewsAsync(siteUrl, cleanId);
                Views = new ObservableCollection<SPViewData>(viewsData);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Ошибка вью: {ex.Message}");
            }



			try
			{
				var allItems = await _spService.GetListItemsByIDAsync(siteUrl, cleanId);

				if (allItems.Count > 250)
				{
					LogAndStatus($"Внимание: в списке {allItems.Count} элементов. Показаны первые 250.");
						Items = new ObservableCollection<SPListItemData>(allItems.Take(250));
				}
				else
				{
					LogAndStatus($"Загружено элементов: {allItems.Count}");
					Items = new ObservableCollection<SPListItemData>(allItems);
				}				
			}
			catch(Exception ex)
			{
               LogAndStatus($"Ошибка загрузки элементов списка: {ex.Message}");
				
			}
        }
    }
}