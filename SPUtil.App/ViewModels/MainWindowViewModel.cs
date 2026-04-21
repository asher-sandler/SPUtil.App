
using SPUtil.App.Views;
using SPUtil.Infrastructure;
using SPUtil.Services;
using System.Collections.ObjectModel;
using System.Windows;




namespace SPUtil.App.ViewModels
{
    public partial class MainWindowViewModel : BindableBase
    {
        private readonly ISharePointService _spService;
        private readonly IContainerExtension _container;

        private string _leftSiteUrl  = string.Empty;
        private string _rightSiteUrl = string.Empty;
		//string.Empty;
        private ObservableCollection<SPNode>? _leftSiteNodes;
        private ObservableCollection<SPNode>? _rightSiteNodes;
        private SPNode? _selectedLeftNode;
        private SPNode? _selectedRightNode;
        
        // Две независимые панели деталей
        private object? _leftDetailsView;
        private object? _rightDetailsView;
        private string? _detailedInfo;

		private string _connectionStatus = "Not Connected";
		private string _currentUserName;

        private readonly SharePointCloneService _cloneService;

        public string ConnectionStatus { get => _connectionStatus; set => SetProperty(ref _connectionStatus, value); }
		public string CurrentUserName { get => _currentUserName; set => SetProperty(ref _currentUserName, value); }

		// В конструкторе
		public SPNode? SelectedLeftNode
		{
			get => _selectedLeftNode;
			set
			{
				if (SetProperty(ref _selectedLeftNode, value))
				{
					// Уведомляем команду, что нужно перепроверить доступность
					CopyEmptyListCommand.RaiseCanExecuteChanged();
				}
			}
		}
		
        public string LeftSiteUrl { get => _leftSiteUrl; set => SetProperty(ref _leftSiteUrl, value); }
        public string RightSiteUrl { get => _rightSiteUrl; set => SetProperty(ref _rightSiteUrl, value); }
        
        public ObservableCollection<SPNode>? LeftSiteNodes { get => _leftSiteNodes; set => SetProperty(ref _leftSiteNodes, value); }
        public ObservableCollection<SPNode>? RightSiteNodes { get => _rightSiteNodes; set => SetProperty(ref _rightSiteNodes, value); }

        public object? LeftDetailsView { get => _leftDetailsView; set => SetProperty(ref _leftDetailsView, value); }
        public object? RightDetailsView { get => _rightDetailsView; set => SetProperty(ref _rightDetailsView, value); }
        
        public string? DetailedInfo { get => _detailedInfo; set => SetProperty(ref _detailedInfo, value); }

        public DelegateCommand ConnectLeftCommand { get; }
        public DelegateCommand ConnectRightCommand { get; }
        public DelegateCommand<SPNode> LeftNodeSelectedCommand { get; }
        public DelegateCommand<SPNode> RightNodeSelectedCommand { get; }

		public DelegateCommand CopyEmptyListCommand { get; }
		public DelegateCommand CopyListWithDataCommand { get; private set; }
		public DelegateCommand DeleteListCommand { get; private set; }
		public DelegateCommand CompareListsCommand { get; private set; }
		public DelegateCommand ExportListCommand { get; private set; }
		public DelegateCommand ShowSchemaCommand { get; }
        public DelegateCommand ExitCommand { get; }

        public DelegateCommand ExportToUniversalWindowCommand { get; }

		private bool _isLeftConnected;
		public bool IsLeftConnected
		{
			get => _isLeftConnected;
			set => SetProperty(ref _isLeftConnected, value);
		}

		private bool _isRightConnected;
		public bool IsRightConnected
		{
			get => _isRightConnected;
			set => SetProperty(ref _isRightConnected, value);
		}
		
		public string LeftSiteFullLink => IsLeftConnected && !string.IsNullOrWhiteSpace(LeftSiteUrl) 
			? $"{SPUsingUtils.UrlWithF5(LeftSiteUrl).TrimEnd('/')}/_layouts/15/viewlsts.aspx" 
			: string.Empty;

		public string RightSiteFullLink => IsRightConnected && !string.IsNullOrWhiteSpace(RightSiteUrl) 
			? $"{SPUsingUtils.UrlWithF5(RightSiteUrl).TrimEnd('/')}/_layouts/15/viewlsts.aspx" 
			: string.Empty;
			
		// В объявлении свойств команд:
		public DelegateCommand ConnectAsCommand { get; }

		// В конструкторе ViewModel:
		
		private string _statusMessage = "Ready";
		public string StatusMessage 
		{ 
			get => _statusMessage; 
			set => SetProperty(ref _statusMessage, value); 
		}


		private bool _isExporting;
		public bool IsExporting
		{
			get => _isExporting;
			set
			{
				_isExporting = value;
				// Уведомляем UI, что нужно показать/скрыть ProgressBar
				RaisePropertyChanged(nameof(IsExporting)); 
			}
		}

		private int _exportProgress;
		public int ExportProgress
		{
			get => _exportProgress;
			set
			{
				_exportProgress = value;
				// Уведомляем UI об изменении %
				RaisePropertyChanged(nameof(ExportProgress));
			}
		}
		
		private string _previewText;
		public string PreviewText { get => _previewText; set => SetProperty(ref _previewText, value); }

		private CancellationTokenSource _exportCts;

		private ObservableCollection<DialogButton> _dialogButtons;
		public ObservableCollection<DialogButton> DialogButtons { get => _dialogButtons; set => SetProperty(ref _dialogButtons, value); }		
        private async void CheckService()
        {
            bool isAlive = await _cloneService.TestCloneServiceConnectionAsync();
            System.Diagnostics.Debug.WriteLine($"Clone Service status: {isAlive}");
        }
        public MainWindowViewModel(ISharePointService spService, SharePointCloneService cloneService, IContainerExtension container)
        //public MainWindowViewModel(ISharePointService spService, IContainerExtension container)
        {
            _spService = spService;
            _container = container;
            _cloneService = cloneService; // Сохраняем его

            // Load site URLs from appsettings.json (excluded from git)
            var (leftUrl, rightUrl) = LoadAppSettings();
            _leftSiteUrl  = leftUrl;
            _rightSiteUrl = rightUrl;

            CheckService();

			ConnectLeftCommand = new DelegateCommand(async () =>
			{
				

				if (IsUrlEmpty(LeftSiteUrl)) return;
				
				IsLeftConnected = false;
				ConnectionStatus = "Connecting...";

				try
				{
					var nodes = await PerformConnectionAsync(LeftSiteUrl);
					if (nodes != null && nodes.Count > 0)
					{
						LeftSiteNodes = nodes;
						IsLeftConnected = true;
						ConnectionStatus = "Connected";
					}
				}
				catch (Exception ex)
				{
					ConnectionStatus = $"Error: {ex.Message}";
				}
				finally
				{
					RaisePropertyChanged(nameof(IsLeftConnected));
					RaisePropertyChanged(nameof(LeftSiteFullLink));
				}
			});

			ConnectRightCommand = new DelegateCommand(async () =>
			{
				if (IsUrlEmpty(RightSiteUrl)) return; 

				IsRightConnected = false;
				ConnectionStatus = "Connecting...";

				try
				{
					var nodes = await PerformConnectionAsync(RightSiteUrl);
					if (nodes != null && nodes.Count > 0)
					{
						RightSiteNodes = nodes;
						IsRightConnected = true;
						ConnectionStatus = "Connected";
					}
				}
				catch (Exception ex)
				{
					ConnectionStatus = $"Error: {ex.Message}";
				}
				finally
				{
					RaisePropertyChanged(nameof(IsRightConnected));
					RaisePropertyChanged(nameof(RightSiteFullLink));
				}
			});
			
			
			ConnectAsCommand = new DelegateCommand(OnConnectAs);

            ExitCommand = new DelegateCommand(OnExit);


            CurrentUserName = _spService.GetCurrentUsername();
            //ConnectionStatus = _spService.GetConnectionStatus();

            // Обработка выбора слева
            LeftNodeSelectedCommand = new DelegateCommand<SPNode>(async node =>
            {
                _selectedLeftNode = node;
                await UpdateDetailsAsync(LeftSiteUrl, node, true);
            });

            // Обработка выбора справа
            RightNodeSelectedCommand = new DelegateCommand<SPNode>(async node =>
            {
                _selectedRightNode = node;
                await UpdateDetailsAsync(RightSiteUrl, node, false);
            });
            CopyEmptyListCommand = new DelegateCommand(ExecuteCopyEmptyList, CanExecuteCopy)
                    .ObservesProperty(() => SelectedLeftNode)
                    .ObservesProperty(() => RightSiteUrl);

           CopyListWithDataCommand = new DelegateCommand(ExecuteCopyListWithData, CanExecuteCopy)
                    .ObservesProperty(() => SelectedLeftNode)
                    .ObservesProperty(() => RightSiteUrl);

            DeleteListCommand = new DelegateCommand(ExecuteDeleteList, CanExecuteCopy)
                    .ObservesProperty(() => SelectedLeftNode)
                    .ObservesProperty(() => RightSiteUrl);

           CompareListsCommand = new DelegateCommand(ExecuteCompareList, CanExecuteCopy)
                    .ObservesProperty(() => SelectedLeftNode)
                    .ObservesProperty(() => RightSiteUrl);

            ExportListCommand = new DelegateCommand(ExecuteExportList, CanExecuteCopy)
                    .ObservesProperty(() => SelectedLeftNode)
                    .ObservesProperty(() => RightSiteUrl);

			ShowSchemaCommand = new DelegateCommand(async () => await ExecuteShowSchema());

             // Подписка на изменение URL для обновления доступности команд
 


			this.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == nameof(RightSiteUrl))
                {
                    CopyEmptyListCommand.RaiseCanExecuteChanged();
                    CopyListWithDataCommand.RaiseCanExecuteChanged();
                }
            };
			
			//ShowSchemaCommand = new DelegateCommand(async () => await ExecuteShowSchema());			
        }
        private void OnExit()
        {
            var result = System.Windows.MessageBox.Show(
                    "Are you sure you want to exit the application?",
                    "Confirm Exit",
                    System.Windows.MessageBoxButton.YesNo,
                    System.Windows.MessageBoxImage.Question);

            
            if (result == System.Windows.MessageBoxResult.Yes)
            {
                System.Windows.Application.Current.Shutdown();
            }
        }

        private async Task<ObservableCollection<SPNode>?> PerformConnectionAsync(string siteUrl)
		{
			string url = SPUsingUtils.NormalizeUrl(siteUrl);
			bool isAuthorized = false;

			while (!isAuthorized)
			{
				// 1. Если в реестре пусто — сразу окно
				if (SPUsingUtils.GetCredentials() == null)
				{
					if (!ShowLoginDialog()) return null;
				}

				ConnectionStatus = "Validating access...";
				var authResult = await _spService.ValidateConnectionAsync(url);

				switch (authResult)
				{
					case AuthResult.Success:
						isAuthorized = true;
						break;

					case AuthResult.InvalidCredentials:
						MessageBox.Show("Invalid username or password.", "Auth Error", MessageBoxButton.OK, MessageBoxImage.Error);
						if (!ShowLoginDialog()) return null;
						break;

					case AuthResult.AccessDenied:
						MessageBox.Show("Access Denied: You don't have permissions for this site.", "Error", MessageBoxButton.OK, MessageBoxImage.Stop);
						ConnectionStatus = "Access Denied";
						return null;

					case AuthResult.SiteNotFound:
						MessageBox.Show("Site not found. Check the URL.", "404", MessageBoxButton.OK, MessageBoxImage.Warning);
						ConnectionStatus = "Wrong site address";
						return null;

					default:
						ConnectionStatus = "Connection error.";
						return null;
				}
			}

			ConnectionStatus = "Loading structure...";
			return await _spService.GetSiteStructureAsync(url);
		}		
        private async Task<ObservableCollection<SPNode>> ConnectWithAuthRetry(string url)
        {
            while (true) // Цикл для повтора при ошибке
            {
                try
                {
                    var nodes = await _spService.GetSiteStructureAsync(url);

                    // Если в сервисе стоит catch { }, и он возвращает пустой список при ошибке, 
                    // нам нужно проверить, не пустой ли он из-за ошибки авторизации.
                    // Но лучше, если сервис прокидывает Exception.

                    if (nodes == null || nodes.Count == 0)
                    {
                        // Если список пуст, возможно пароль неверный. 
                        // Для простоты вызовем логин, если реестр пуст или была ошибка.
                    }

                    return nodes;
                }
                catch (Exception ex)
                {
                    // Ошибка авторизации (401) или доступа
                    var result = ShowLoginDialog();
                    if (!result) return new ObservableCollection<SPNode>(); // Пользователь отменил

                    // Если сохранили новый пароль - цикл while попробует снова
                }
            }
        }
		// Сам метод реализации:
		private void OnConnectAs()
		{
			// Мы просто вызываем созданный ранее метод. 
			// Если пользователь введет данные и нажмет Save, реестр обновится.
			if (ShowLoginDialog())
			{
				// После смены пользователя можно обновить имя в статусбаре
				// Предполагается, что у вас есть свойство CurrentUserName
				RaisePropertyChanged(nameof(CurrentUserName)); 
				
				ConnectionStatus = "Credentials updated. Please reconnect to sites.";
			}
		}		

        private bool ShowLoginDialog()
        {
            bool isSaved = false;
            // Используем Dispatcher, чтобы окно открылось в UI потоке
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                var loginWin = new LoginWindow { Owner = System.Windows.Application.Current.MainWindow };
                if (loginWin.ShowDialog() == true)
                {
                    // Сохраняем в реестр
                    SPUsingUtils.SaveCredentials(loginWin.UserName, loginWin.Password);
                    isSaved = true;
                }
            });
            return isSaved;
        }
		private bool IsUrlEmpty(string url){
				bool retValue = true;
				if (string.IsNullOrWhiteSpace(url))
				{
					System.Windows.MessageBox.Show(
						"Please enter the Site URL before connecting.", 
						"Connection Error", 
						System.Windows.MessageBoxButton.OK, 
						System.Windows.MessageBoxImage.Warning);
					
					StatusMessage = "Error: URL is empty";
					
				}
				else
				{
					retValue = false;
				}					
			return retValue;
		}
        //
        private async Task ExecuteShowSchema()
		{
			if (SelectedLeftNode == null || SelectedLeftNode.Type != SharePointObjectType.List)
			{
				System.Windows.MessageBox.Show("Select a list in the left panel.");
				return;
			}
			try 
			{
				StatusMessage = "Load schema...";

				// ШАГ 1: Получаем "грязный" список XML из SharePoint (через сервис)
				// Это асинхронная операция (Task), так как идем в сеть
				string listTitle = SelectedLeftNode.Title;

                var schemasHead1 = new List<string>();
                var schemasHead2 = new List<string>();
                schemasHead1.Add($"<!-- -->");
                schemasHead1.Add($"<!-- List '{listTitle}' Fields Schema -->");
                schemasHead1.Add($"<!-- -->");

                schemasHead2.Add($"<!-- -->");
                schemasHead2.Add($"<!-- List '{listTitle}' Views Schema -->");
                schemasHead2.Add($"<!-- -->");



                List<string> rawSchemas = await _spService.GetListSchemaAsync(LeftSiteUrl, listTitle);
                List<string> viewSchemas = await _spService.GetListViewSchemasAsync(LeftSiteUrl, listTitle);

                // ШАГ 2: Очищаем каждое поле
                // Это синхронная операция (string), выполняется в памяти
                // Используем LINQ для удобства
                var cleanSchemas = rawSchemas
					.Select(xml => _cloneService.CleanFieldXml(xml))
					.ToList();

				// ШАГ 3: Подготовка текста для окна
				/*
				string finalXml = string.Join(Environment.NewLine + "" + Environment.NewLine, cleanSchemas);

				// ШАГ 4: Вызов окна
				var viewer = new SPUtil.App.Views.SchemaViewerWindow(finalXml)
				{
					Owner = System.Windows.Application.Current.MainWindow
				};
				viewer.ShowDialog();
				*/
				string cleanFieldSchemas = string.Join(Environment.NewLine, cleanSchemas);


                var finalLines = new List<string>();

                finalLines.AddRange(schemasHead1);     // Заголовок полей
                finalLines.AddRange(cleanFieldSchemas); // Сами поля
                finalLines.AddRange(schemasHead2);     // Заголовок вьюх
                finalLines.AddRange(viewSchemas);      // Сами вьюхи

                // 5. Формируем финальную строку
                string resultXml = string.Join(Environment.NewLine, finalLines);
                var viewer = new SPUtil.App.Views.SchemaViewerWindow(SPUsingUtils.FormatXml(resultXml), listTitle)
                {
                    Owner = System.Windows.Application.Current.MainWindow,
                    // Центрируем относительно родительского окна
                    WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
                    // Указываем, что окно всегда должно быть сверху своего владельца
                    Topmost = false // Оставляем false, чтобы не перекрывать другие приложения, 
                                    // но благодаря ShowDialog оно будет модальным внутри приложения
                };

                // Подписываемся на событие загрузки, чтобы "дернуть" фокус на себя
                viewer.Loaded += (s, e) =>
                {
                    viewer.Activate();
                    viewer.Focus();
                };

                viewer.ShowDialog();

                StatusMessage = "Ready";
            }
			catch (Exception ex)
			{
				System.Windows.MessageBox.Show(ex.Message);
			}
		}
		//		
		private async Task LoadSiteAsync(bool isLeft)
		{
			// Устанавливаем статус начала работы
			ConnectionStatus = isLeft ? "Connecting to Source..." : "Connecting to Target...";
			
			try 
			{
				var url = isLeft ? LeftSiteUrl : RightSiteUrl;
				if (string.IsNullOrEmpty(url)) 
				{
					ConnectionStatus = "Error: URL is empty";
					return;
				}

				var nodes = await _spService.GetSiteStructureAsync(url);
				
				if (isLeft) LeftSiteNodes = nodes; 
				else RightSiteNodes = nodes;

				// Устанавливаем статус успеха
				//ConnectionStatus = "Connected successfully";
				ConnectionStatus = "Connected successfully"; 
			}
			catch (Exception ex) 
			{
				// Выводим ошибку в статус-бар
				//ConnectionStatus = $"Connection Error: {ex.Message}";
				ConnectionStatus = $"Error: {ex.Message}"; 
			}
		}
        private async Task UpdateDetailsAsync(string siteUrl, SPNode? node, bool isLeftPane)
        {
            // Сбрасываем текущую панель перед загрузкой
            if (isLeftPane) LeftDetailsView = null; else RightDetailsView = null;

            if (node == null) return;

            object? newView = null;

            // Проверяем, является ли узел списком/библиотекой (у них есть Tag с ID шаблона)
            if (int.TryParse(node.Tag, out int templateId))
            {
                try 
                {
                    if (templateId == 100) // Обычный список
                    {
                        var vm = _container.Resolve<List100ViewModel>();
                        vm.ListTitle = node.Title;
						vm.IsSourceMode = isLeftPane; // true для левой, false для правой
                        await vm.LoadDataAsync(siteUrl, node.Path);
                        newView = vm;
                    }
                    else if (templateId == 101) // Библиотека документов
                    {
                        var vm = _container.Resolve<Library101ViewModel>();
                        vm.LibraryTitle = node.Title;
						vm.IsSourceMode = isLeftPane;
                        await vm.LoadDataAsync(siteUrl, node.Path);
                        newView = vm;
                    }
                    else if (templateId == 850 || templateId == 119) // Страницы (Site Pages / Wiki)
                    {

						var vm = _container.Resolve<PagesViewModel>();
						vm.IsSourceMode = isLeftPane;

						// Give the VM access to the target site URL so CopyPageCommand knows
						// where to copy. Source (left) pane targets the right site, and vice-versa.
						// isLeftPane=true  → source is left  → target is RightSiteUrl
						// isLeftPane=false → source is right → target is LeftSiteUrl  (less common)
						vm.SetTargetSiteUrl(isLeftPane
							? SPUtil.Infrastructure.SPUsingUtils.NormalizeUrl(RightSiteUrl)
							: SPUtil.Infrastructure.SPUsingUtils.NormalizeUrl(LeftSiteUrl));

						await vm.LoadDataAsync(siteUrl, node.Path);
						newView = vm;


                    }
                    else 
                    {
                        // Для всех остальных типов (настройки, папки и т.д.)
                        newView = new InfoViewModel($"Тип: {node.Type}\nID шаблона: {templateId}\nИмя: {node.Title}\n\nЭтот тип объекта пока не имеет специального интерфейса.");
                    }
                }
                catch (Exception ex)
                {
                    newView = new InfoViewModel($"Ошибка загрузки деталей: {ex.Message}");
                }
            }
            else
            {
                // Если кликнули на узел без ID шаблона (например, корень сайта или системную папку)
                newView = new InfoViewModel($"Объект: {node.Title}\nТип: {node.Type}\nПуть: {node.Path}");
            }

            // Назначаем созданную вью-модель в нужную панель
            if (isLeftPane) LeftDetailsView = newView; else RightDetailsView = newView;
        }
		private async void ExecuteCopyEmptyList()
		{
			await StartCopyProcess(withData: false);
		}
		
		private async void ExecuteCopyListWithData()
		{
			await StartCopyProcess(withData: true);
		}		
		//
		private async void ExecuteDeleteList()
		{
			await DeleteList();
		}		
		//
		private async void ExecuteCompareList()
		{
			await CompareList();
		}		
		//
		private async void ExecuteExportList()
		{
			await ExportList();
		}		
		//
		private async Task DeleteList()
		{
			if (SelectedLeftNode == null || SelectedLeftNode.Type != SharePointObjectType.List)
			{
				System.Windows.MessageBox.Show("Select a list in the left panel.");
				return;
			}

			string listTitle = SelectedLeftNode.Title;
			StatusMessage = "Delete list...";

			
			var result = System.Windows.MessageBox.Show(
				$"Delete list:\n\n" +
				$"List: {listTitle}\n" +
				$"Site: {LeftSiteUrl}\n" +
				$"Continue?", 
				"Confirm Delete", 
				System.Windows.MessageBoxButton.YesNo, 
				System.Windows.MessageBoxImage.Question);

			if (result == System.Windows.MessageBoxResult.Yes)
			{
				StatusMessage = "Delete list '"+listTitle+"' in progress...";
				// Здесь будет вызов метода Delete списка из сервиса
			}
			
			
		}
		private async Task CompareList()
		{
			if (SelectedLeftNode == null || SelectedLeftNode.Type != SharePointObjectType.List)
			{
				System.Windows.MessageBox.Show("Select a list in the left panel.");
				return;
			}

			string listTitle = SelectedLeftNode.Title;
			string targetUrl = SPUsingUtils.NormalizeUrl(RightSiteUrl);

			bool exists = await _spService.ListExistsAsync(targetUrl, listTitle);
			if (!exists)
			{
				System.Windows.MessageBox.Show($"List '{listTitle}' not exists on destination site!", "Info");
				return;
			}

			// Сразу готовим кнопки для окна
			DialogButtons = new ObservableCollection<DialogButton>
			{

				new DialogButton 
				{ 
					Caption = "Copy", 
					Action = async () => {
						if (string.IsNullOrEmpty(PreviewText) || PreviewText.StartsWith("Comparing")) return;
						
						System.Windows.Clipboard.SetText(PreviewText);
						string oldStatus = StatusMessage;
						StatusMessage = "Copied to clipboard!";
						
						await Task.Delay(5000);
						if (StatusMessage == "Copied to clipboard!")
							StatusMessage = "Ready";
					}
				},				
				new DialogButton { 
					Caption = "Close", 
					IsCancel = true, 
					Action = () => {

							var winToClose = System.Windows.Application.Current.Windows
								.OfType<SPUtil.App.Views.UniversalPreviewWindow>()
								.FirstOrDefault(w => w.DataContext == this);

							winToClose?.Close();
						} 
					}
				
			};

			var win = new SPUtil.App.Views.UniversalPreviewWindow()
			{
				Title = $"Compare: {listTitle}",
				Owner = System.Windows.Application.Current.MainWindow,
				DataContext = this
			};

			// Запускаем сравнение сразу после загрузки окна
			win.Loaded += async (s, e) =>
			{
				StatusMessage = "Fetching schemas...";
				PreviewText = "Comparing... Please wait.";
				
				try 
				{
					// Получаем сырые XML схем с обоих сайтов
					var leftSchemas = await _spService.GetListSchemaAsync(LeftSiteUrl, listTitle);
					var rightSchemas = await _spService.GetListSchemaAsync(RightSiteUrl, listTitle);

					// Выполняем сравнение в фоновом потоке
					string report = await Task.Run(() => CompareTwoLists(leftSchemas, rightSchemas,LeftSiteUrl,RightSiteUrl));
					
					PreviewText = report;
					StatusMessage = "Compare complete.";
				}
				catch (Exception ex)
				{
					PreviewText = "Error during compare: " + ex.Message;
					StatusMessage = "Error";
				}
			};

			win.ShowDialog();
		}
		private async Task<string> CompareTwoLists(List<string> leftFieldsXml, List<string> rightFieldsXml, string leftSiteUrl, string rightSiteUrl)
		{
			string leftListTitle = SelectedLeftNode.Title;
			var sb = new System.Text.StringBuilder();
			sb.AppendLine($"--- Comparison Report ---");
			sb.AppendLine($"Source URL: {leftSiteUrl}");
			sb.AppendLine($"Target URL: {rightSiteUrl}");
			sb.AppendLine($"List: {leftListTitle}");
			sb.AppendLine($"Generated: {DateTime.Now:yyyy-MM-dd HH:mm:ssK}");
			sb.AppendLine();

			var leftDict = ParseFieldsToDictionary(leftFieldsXml);
			var rightDict = ParseFieldsToDictionary(rightFieldsXml);

			// ... (Блоки MISSING и EXTRA оставляем без изменений) ...

			sb.AppendLine("[ SCHEMA DIFFERENCES ]");
			bool hasDiffs = false;
			var commonFields = leftDict.Keys.Intersect(rightDict.Keys).ToList();

			foreach (var name in commonFields)
			{
				string leftRaw = leftDict[name];
				string rightRaw = rightDict[name];
				
				string leftClean = _cloneService.CompareCleanFieldXml(leftRaw);
				string rightClean = _cloneService.CompareCleanFieldXml(rightRaw);

				// Если строки очищенного XML не равны
				if (!string.Equals(leftClean, rightClean, StringComparison.OrdinalIgnoreCase))
				{
					// --- ЛОГИКА ДЛЯ LOOKUP ---
					if (GetFieldType(leftRaw) == "Lookup" && GetFieldType(rightRaw) == "Lookup")
					{
						string leftListId = GetAttributeFromXml(leftRaw, "List");
						string rightListId = GetAttributeFromXml(rightRaw, "List");

						// Получаем реальные имена списков
						string leftListName = await _spService.GetListNameByIdAsync(leftSiteUrl, leftListId);
						string rightListName = await _spService.GetListNameByIdAsync(rightSiteUrl, rightListId);

						// Если имена списков совпали — игнорируем разницу в XML (в GUID-ах)
						if (leftListName.Equals(rightListName, StringComparison.OrdinalIgnoreCase))
						{
							continue; 
						}

						// Если имена разные, выводим разницу и дописываем имена списков для наглядности
						hasDiffs = true;
						sb.AppendLine($"FIELD: {name} (Lookup target mismatch)");
						sb.AppendLine("  LEFT (Source):");
						sb.AppendLine(IndentText(leftClean, 4));
						sb.AppendLine($"    Lookup List: {leftListName}");
						sb.AppendLine("  RIGHT (Dest):");
						sb.AppendLine(IndentText(rightClean, 4));
						sb.AppendLine($"    Lookup List: {rightListName}");
					}
					else
					{
						// Для всех остальных типов полей просто выводим разницу
						hasDiffs = true;
						sb.AppendLine($"FIELD: {name}");
						sb.AppendLine("  LEFT (Source):");
						sb.AppendLine(IndentText(leftClean, 4));
						sb.AppendLine("  RIGHT (Dest):");
						sb.AppendLine(IndentText(rightClean, 4));
					}

					sb.AppendLine(new string('-', 30));
				}
			}

			if (!hasDiffs) sb.AppendLine("All common fields have identical schemas (including Lookup names mapping).");
			return sb.ToString();
		}
		// Вспомогательный метод для извлечения атрибутов из сырого XML
		private string GetAttributeFromXml(string xml, string attrName)
		{
			try {
				var el = System.Xml.Linq.XElement.Parse(xml);
				return el.Attribute(attrName)?.Value;
			} catch { return null; }
		}
		// Вспомогательный метод: создаем словарь [InternalName] -> [SchemaXml]
		private Dictionary<string, string> ParseFieldsToDictionary(List<string> xmls)
		{
			var dict = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
			foreach (var xml in xmls)
			{
				try {
					var el = System.Xml.Linq.XElement.Parse(xml);
					string name = el.Attribute("Name")?.Value;
					if (!string.IsNullOrEmpty(name)) dict[name] = xml;
				} catch { /* игнорируем битый XML */ }
			}
			return dict;
		}

		private string GetFieldType(string xml) => System.Xml.Linq.XElement.Parse(xml).Attribute("Type")?.Value ?? "Unknown";

		private string IndentText(string text, int spaces)
		{
			string indent = new string(' ', spaces);
			return indent + text.Replace(Environment.NewLine, Environment.NewLine + indent);
		}		
		private async Task ExportList()
		{
			if (SelectedLeftNode == null || SelectedLeftNode.Type != SharePointObjectType.List)
			{
				System.Windows.MessageBox.Show("Select a list in the left panel.");
				return;
			}
			string listTitle = SelectedLeftNode.Title;
			StatusMessage = "Export list...";



				StatusMessage = "Export list '"+listTitle+"'.";
				await ExportToUniversalWindow(listTitle);
				
			//}
			
			
		}

	
        private async Task ExportToUniversalWindow(string listTitle)
		{
			_exportCts?.Cancel(); // Останавливаем старые задачи, если они были
			IsExporting = false;
			ExportProgress = 0;
			PreviewText = "Select export mode...";
			StatusMessage = "Ready";			
			// Инициализируем источник токена отмены
			_exportCts = new CancellationTokenSource();

			DialogButtons = new ObservableCollection<DialogButton>
			{
				new DialogButton 
				{ 
					Caption = "SQL", 
					Action = async () => {
						if (IsExporting ) return;
						IsExporting  = true;
						try {
							var progressHandler = new Progress<int>(p => ExportProgress = p);
							StatusMessage = "Gen SQL Script...";
							PreviewText = "Started. Please wait...";
							
							// Передаем токен в Task.Run (если методы генерации поддерживают отмену)
							var result = await Task.Run(() => GenerateSqlScript(listTitle, progressHandler,_exportCts.Token), _exportCts.Token);
							
							PreviewText = result;
							StatusMessage = "SQL Script is ready.";
						}
						catch (OperationCanceledException) {
							StatusMessage = "Export cancelled.";
						}
						catch (Exception ex) {
							StatusMessage = "Error: " + ex.Message;
						}
						finally {
							IsExporting  = false;
							ExportProgress = 0;
						}
					}
				},
				new DialogButton 
				{ 
					Caption = "CSV", 
					Action = async () => {
						if (IsExporting ) return;
						IsExporting  = true;
						try {
							var progressHandler = new Progress<int>(p => ExportProgress = p);
							StatusMessage = "Gen CSV...";
							PreviewText = "Started. Please wait...";

							var result = await Task.Run(() => GenerateCsvData(listTitle, progressHandler, _exportCts.Token), _exportCts.Token);
							
							PreviewText = result;
							StatusMessage = "CSV is ready. After insert data to Excel press ALT + A + E";
						}
						catch (OperationCanceledException) {
							StatusMessage = "Export cancelled.";
						}
						finally {
							IsExporting  = false;
							ExportProgress = 0;
						}
					}
				},
				new DialogButton 
				{ 
					Caption = "Copy", 
					Action = async () => {
						if (string.IsNullOrEmpty(PreviewText) || PreviewText.StartsWith("Started")) return;
						
						System.Windows.Clipboard.SetText(PreviewText);
						string oldStatus = StatusMessage;
						StatusMessage = "Copied to clipboard!";
						
						await Task.Delay(5000);
						if (StatusMessage == "Copied to clipboard!")
							StatusMessage = "Ready";
					}
				},
				new DialogButton 
				{ 
					Caption = "Close", 
					IsCancel = true, 
					Action = () => {
						// Пытаемся остановить фоновые потоки
						_exportCts?.Cancel();

						var winToClose = System.Windows.Application.Current.Windows
							.OfType<SPUtil.App.Views.UniversalPreviewWindow>()
							.FirstOrDefault(w => w.DataContext == this);

						winToClose?.Close();
					}
				}
			};

			var win = new SPUtil.App.Views.UniversalPreviewWindow()
			{
				Title = $"Export: {listTitle}",
				Owner = System.Windows.Application.Current.MainWindow,
				WindowStartupLocation = System.Windows.WindowStartupLocation.CenterOwner,
				DataContext = this
			};

			// Если пользователь закроет окно "крестиком", тоже отменяем поток
			win.Closing += (s, e) => {
				_exportCts?.Cancel();
			};

			win.Loaded += (s, e) => {
				win.Activate();
				win.Focus();
			};

			win.ShowDialog();
		}
//		

		private async Task<string> GenerateSqlScript(string listTitle, IProgress<int> progress, CancellationToken ct)
		{
			// Если внутри метода экспорта:

			var rawSchemas = await _spService.GetListSchemaAsync(LeftSiteUrl, listTitle);

			var progressHandler = new Progress<int>(p => ExportProgress = p);
			var items = await _spService.GetListItemsByTitleAsync(LeftSiteUrl, 
								listTitle, 
								progressHandler, 
								ct  
							);

			int total = items.Count;
			var sb = new System.Text.StringBuilder();
			string tableName = listTitle.Replace(" ", "_");

			// --- Часть 1: CREATE TABLE ---
			sb.AppendLine("/*");
			sb.AppendLine($"	Database migration script for SharePoint list: {listTitle}");
			sb.AppendLine($"	Generated on {DateTime.Now:yyyy-MM-dd HH:mm}");
			sb.AppendLine("*/");
			sb.AppendLine("");
			
			
			sb.AppendLine($"IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[{tableName}]') AND type in (N'U'))");
			sb.AppendLine("BEGIN");
			sb.AppendLine($"    CREATE TABLE [dbo].[{tableName}] (");
			
			// 1. Первичный ключ SQL (автоинкремент)
			sb.AppendLine("        [ID] INT IDENTITY(1,1) PRIMARY KEY,");
			
			// 2. Поле для хранения оригинального ID из SharePoint
			sb.AppendLine("        [SPID] INT NULL,");

			var columns = new List<string>();
			var internalNames = new List<string>();

			foreach (var xml in rawSchemas)
			{
				var el = System.Xml.Linq.XElement.Parse(xml);
				string name = el.Attribute("Name")?.Value;
				string type = el.Attribute("Type")?.Value;

				if (string.IsNullOrEmpty(name)) continue;
				
				// Пропускаем ID, так как мы его уже добавили вручную как SPID
				if (name.Equals("ID", StringComparison.OrdinalIgnoreCase))
				{
					internalNames.Add(name); // Добавляем в список для маппинга данных
					continue; 
				}
				
				internalNames.Add(name); 
				string sqlType = GetSqlType(type);
				columns.Add($"        [{name}] {sqlType} NULL");
			}

			sb.AppendLine(string.Join("," + Environment.NewLine, columns));
			sb.AppendLine("    );");
			
			// 3. Создаем НЕ уникальный индекс для SPID
			sb.AppendLine();
			sb.AppendLine($"    CREATE INDEX [IX_{tableName}_SPID] ON [dbo].[{tableName}] ([SPID]);");
			
			sb.AppendLine("END");
			sb.AppendLine("GO");
			sb.AppendLine();

			// --- Часть 2: INSERT INTO ---
			sb.AppendLine($"-- Insert {items.Count} row(s)");
			
			// В данном случае IDENTITY_INSERT НЕ НУЖЕН, так как мы не трогаем колонку [ID]
			// SQL сам будет генерировать 1, 2, 3... в колонке ID
			// А мы будем вставлять оригинальные значения в SPID

				for (int i = 0; i < total; i++)
    			{
				var item = items[i];	
				// Формируем список имен колонок (заменяем ID на SPID)
				var colNames = string.Join(", ", internalNames.Select(n => 
					n.Equals("ID", StringComparison.OrdinalIgnoreCase) ? "[SPID]" : $"[{n}]"));

				var values = internalNames.Select(col => 
				{
					object val = item.Values.ContainsKey(col) ? item.Values[col] : null;
					return SqlEscape(val); 
				});

				sb.AppendLine($"INSERT INTO [dbo].[{tableName}] ({colNames}) VALUES ({string.Join(", ", values)});");
				// Каждые 50 строк обновляем статус-бар, чтобы не перегружать UI
				if (i % 50 == 0 || i == total - 1)
				{
					int percent = (int)((i + 1) * 100.0 / total);
					progress.Report(percent);
				}				
			}
			
			sb.AppendLine("GO");

			return sb.ToString();
		}
		
		private string SqlEscape(object val)
		{
			if (val == null || val == DBNull.Value) return "NULL";
			
			// ПРОВЕРКА НА ДАТУ: Самый важный момент
			if (val is DateTime dt)
			{
				// Формат 'YYYY-MM-DDTHH:mm:ss' понимает любой SQL Server
				return $"'{dt.ToString("yyyy-MM-ddTHH:mm:ss")}'";
			}

			string s = "";
			if (val is Microsoft.SharePoint.Client.FieldLookupValue lv) s = lv.LookupValue;
			else if (val is Microsoft.SharePoint.Client.FieldUserValue uv) s = uv.LookupValue;
			else s = val.ToString();

			// Для строк оставляем префикс N и экранирование кавычек
			return $"N'{s.Replace("'", "''")}'";
		}

		private async Task<string> GenerateCsvData(string listTitle, IProgress<int> progress, CancellationToken ct)
		{
			// 1. Получаем схемы (XML)
			// Передаем ct в сервис, чтобы прервать сетевой запрос при закрытии окна
			var rawSchemas = await _spService.GetListSchemaAsync(LeftSiteUrl, listTitle);

			// 2. Получаем данные (Items)
			// progress здесь уже передается извне, используем его напрямую
			var items = await _spService.GetListItemsByTitleAsync(
				LeftSiteUrl, 
				listTitle, 
				progress, 
				ct
			);

			// Подготовка списка колонок
			var columns = rawSchemas.Select(xml => {
				try {
					var el = System.Xml.Linq.XElement.Parse(xml);
					return el.Attribute("Name")?.Value;
				} catch { return null; }
			}).Where(n => !string.IsNullOrEmpty(n)).ToList();

			// 3. Генерация текста в фоновом потоке
			return await Task.Run(() => 
			{
				var sb = new System.Text.StringBuilder();
				int total = items.Count;

				// Header
				sb.AppendLine(string.Join(",", columns.Select(c => CsvEscape(c))));

				// Rows
				for (int i = 0; i < total; i++)
				{
					// ПРОВЕРКА ОТМЕНЫ: Если нажали Close, выбрасываем исключение и выходим
					ct.ThrowIfCancellationRequested();

					var item = items[i];
					var values = columns.Select(col => {
						// Безопасно берем значение из Dictionary FieldValues по InternalName
						object val = item.Values.ContainsKey(col) ? item.Values[col] : "";
						return CsvEscape(val);
					});

					sb.AppendLine(string.Join(",", values));

					// Обновляем прогресс этапа "Генерация CSV" (опционально)
					// Чтобы не путать с загрузкой, можно выводить статус в лог или просто не частить
					if (i % 100 == 0 || i == total - 1)
					{
						int percent = (int)((i + 1) * 100.0 / total);
						progress.Report(percent);
					}
				}

				return sb.ToString();

			}, ct); // Передаем токен и в сам Task.Run
		}

		private string GetSqlType(string spType)
		{
			// Маппинг типов данных SharePoint в типы данных MS SQL
			return spType switch
			{
				"Counter" => "INT",
				"Integer" => "INT",
				"Number"  => "DECIMAL(18,2)",
				"DateTime"=> "DATETIME",
				"Boolean" => "BIT",
				"Note"    => "NVARCHAR(MAX)",
				"User"    => "NVARCHAR(255)", 
				"Lookup"  => "NVARCHAR(255)",
				_         => "NVARCHAR(255)"  // Значение по умолчанию для Text и прочих
			};
		}		
		// Вспомогательный метод для экранирования CSV (чтобы запятые внутри текста не ломали структуру)
		private string CsvEscape(object val)
		{
			if (val == null) return "";
			
			string s = "";

			if (val is DateTime dt)
			{
				// Формат 'YYYY-MM-DD HH:mm:ss' понимает любой Excel
				return $"{dt.ToString("yyyy-MM-dd HH:mm:ss")}";
			}


			// Обработка специфических типов SharePoint
			if (val is Microsoft.SharePoint.Client.FieldLookupValue lv) s = lv.LookupValue;
			else if (val is Microsoft.SharePoint.Client.FieldUserValue uv) s = uv.LookupValue;
			else s = val.ToString();

			// Если в тексте есть запятая, кавычка или перенос строки — оборачиваем в кавычки
			if (s.Contains(",") || s.Contains("\"") || s.Contains("\n") || s.Contains("\r"))
			{
				return $"\"{s.Replace("\"", "\"\"")}\"";
			}
			
			return s;
		}
		/*
		private bool CanExecuteCopy()
		{
			// Проверяем, что выбран именно список
			if (SelectedLeftNode == null || SelectedLeftNode.Type != SharePointObjectType.List)
				return false;

			// Проверяем наличие целевого URL
			if (string.IsNullOrWhiteSpace(RightSiteUrl))
				return false;

			// Проверяем, что это не один и тот же сайт
			string source = _spService.NormalizeUrl(LeftSiteUrl);
			string target = _spService.NormalizeUrl(RightSiteUrl);
			
			return !source.Equals(target, StringComparison.OrdinalIgnoreCase);
		}
*/
		
		private bool CanExecuteCopy() => true;

        // ── Settings loader ───────────────────────────────────────────────────
        /// <summary>
        /// Reads LeftSiteUrl and RightSiteUrl from appsettings.json next to the exe.
        /// The file is excluded from git via .gitignore — copy appsettings.example.json
        /// to appsettings.json and fill in your real URLs.
        /// </summary>
        private static (string leftUrl, string rightUrl) LoadAppSettings()
        {
            try
            {
                string path = System.IO.Path.Combine(
                    AppDomain.CurrentDomain.BaseDirectory, "appsettings.json");

                if (!System.IO.File.Exists(path))
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[AppSettings] appsettings.json not found — URLs left empty. " +
                        "Copy appsettings.example.json → appsettings.json and fill in your URLs.");
                    return (string.Empty, string.Empty);
                }

                string json  = System.IO.File.ReadAllText(path);
                string left  = ExtractJsonString(json, "LeftSiteUrl");
                string right = ExtractJsonString(json, "RightSiteUrl");
                return (left, right);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[AppSettings] Failed to load: {ex.Message}");
                return (string.Empty, string.Empty);
            }
        }

        /// <summary>Extracts a string value from a simple flat JSON object by key name.</summary>
        private static string ExtractJsonString(string json, string key)
        {
            var m = System.Text.RegularExpressions.Regex.Match(
                json,
                $@"""{System.Text.RegularExpressions.Regex.Escape(key)}""\s*:\s*""([^""]*)""");
            return m.Success ? m.Groups[1].Value : string.Empty;
        }
    }
}