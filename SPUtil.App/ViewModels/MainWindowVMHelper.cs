using Microsoft.SharePoint.Portal.WebControls.WSRPWebService;
using SPUtil.Infrastructure;
using System;
using System.Collections.Generic;
using System.Text;
using System.Windows;

namespace SPUtil.App.ViewModels
{
    public partial class MainWindowViewModel : BindableBase
    {
        private async Task StartCopyProcess(bool withData)
        {
            if (SelectedLeftNode == null || SelectedLeftNode.Type != SharePointObjectType.List)
            {
                System.Windows.MessageBox.Show("Выберите список или библиотеку в левой панели.");
                return;
            }

            string sourceTitle = SelectedLeftNode.Title;
            string sourceUrl = SPUsingUtils.NormalizeUrl(LeftSiteUrl);
            string targetUrl = SPUsingUtils.NormalizeUrl(RightSiteUrl);

            StatusMessage = "Loading Info...";
            SPListInfo info = await _spService.GetListDetailedInfoAsync(sourceUrl, sourceTitle);

            //Info
            //	Display Name: Countries
            //	Internal Name: CountriesList
            //	Type: List
            //	Items: 5
            //	Created: 07/07/2022
            // Распределяем логику в зависимости от типа объекта

            if (info.BaseTemplate == 100)
            {
                int templateId = 100;
                await ProcessListCopyAsync(info, targetUrl,  withData, templateId);
            }
            else
            {
                int templateId = 101;
                await ProcessDocLibCopyAsync(info, targetUrl, withData, templateId);

            }
            IsRightConnected = false;
            IsRightConnected = true;
            RaisePropertyChanged(nameof(RightSiteFullLink));
            RaisePropertyChanged(nameof(IsRightConnected));
            
           

        }
        private async Task ProcessListCopyAsync(SPListInfo info, string targetUrl,  bool withData, int templateId)
        {
            // --- ШАГ 1: Первичный диалог (Выбор имени) ---
            var copyDialog = new SPUtil.Views.CopyListDialog(info.Title, info.URL, info.ToString())
            {
                Owner = System.Windows.Application.Current.MainWindow
            };

            if (copyDialog.ShowDialog() != true) return;

            string targetListName = copyDialog.TargetListTitle;
            string action = "Append";
            bool listExists = await _spService.ListExistsAsync(targetUrl, targetListName);

            var infoWin = new SPUtil.Views.OperationInfoWindow
            {
                Owner = System.Windows.Application.Current.MainWindow
            };

            // --- ШАГ 2: Обработка существующего или создание нового списка ---
            if (listExists)
            {
                var existsDialog = new SPUtil.Views.ExistsActionDialog(targetListName);
                if (existsDialog.ShowDialog() != true) return;

                action = existsDialog.SelectedAction; // "Overwrite", "Append" или "Cancel"

                if (action == "Overwrite")
                {
                    StatusMessage = "Удаление старого списка...";
                    infoWin.Show();
                    infoWin.UpdateMessage($"Deleting existing list: {targetListName}...");
                    await _spService.DeleteListAsync(targetUrl, targetListName);

                    infoWin.UpdateMessage($"Creating structure: {targetListName}...");
                    bool created = await CreateListStructureAsync(info, targetUrl, targetListName, templateId);
                    infoWin.Close();
                    if (!created) return;
                }
                else if (action == "Append")
                {
                    StatusMessage = "Будет выполнено добавление в существующий список.";
                }
                else return;
            }
            else
            {
                infoWin.Show();
                infoWin.UpdateMessage($"Creating new list: {targetListName}...");
                bool created = await CreateListStructureAsync(info, targetUrl, targetListName, templateId);
                infoWin.Close();
                if (!created) return;
            }

            // --- ШАГ 3: Копирование данных ---
            if (withData)
            {
                await ExecuteDataCopyAsync(info.URL, targetUrl, info.Title, targetListName, action);
            }

            // --- ШАГ 4: Финализация ---
            RightSiteNodes = await _spService.GetSiteStructureAsync(targetUrl);
            StatusMessage = "Готово";
        }
		private async Task ExecuteDataCopyAsync(string sourceUrl, string targetUrl, string sourceTitle, string targetListName, string action)
		{
			var cts = new CancellationTokenSource();
			var progressWin = new SPUtil.Views.ProgressWindow(cts) { Owner = System.Windows.Application.Current.MainWindow };

			var progressIndicator = new Progress<CopyProgressArgs>(e => {
				progressWin.UpdateStatus(e.Processed, e.Total, e.Message);
			});

			try
			{
				progressWin.Show();
				await _spService.CopyListItemsAsync(sourceUrl, targetUrl, sourceTitle, targetListName, action, progressIndicator, cts.Token);
				progressWin.Close();
			}
			catch (OperationCanceledException)
			{
				System.Windows.MessageBox.Show("Копирование данных отменено.");
			}
			catch (Exception ex)
			{
				progressWin.Close();
				System.Windows.MessageBox.Show($"Ошибка при копировании данных: {ex.Message}");
			}
		}
		
        private static readonly HashSet<string> ReservedLibNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            // Publishing & Media
            "PublishingImages",      // תמונות (Images) - Обязательно PublishingImages в URL!
            "Pages",                 // דפים (Pages) - Для классических сайтов публикации
            "SiteCollectionImages",  // תמונות אוסף אתרים
    
            // Structure & Modern UI
            "SitePages",             // דפי אתר (Site Pages) - Все современные страницы здесь
            "SiteAssets",            // נכסי אתר (Site Assets) - Логотипы, OneNote, скрипты
            "Style Library",         // ספריית סגנונות - CSS и JS оформления
    
            // Core & Documents
            "Shared Documents",      // מסמכים משותפים - Основная библиотека (Documents)
            "Documents",             // Старый стандарт (URL: .../Documents/)
            "SiteCollectionDocuments", 
    
            // Technical
            "FormServerTemplates",   // Шаблоны InfoPath
            "WorkflowTasks"          // Системные задачи рабочих процессов
        };

        private async Task ProcessDocLibCopyAsync(SPListInfo info, string targetUrl, bool withData, int templateId)
        {
            // Step 1: Name selection dialog
            var copyDialog = new SPUtil.Views.CopyListDialog(info.Title, targetUrl, info.ToString()) { Owner = Application.Current.MainWindow };
            if (copyDialog.ShowDialog() != true) return;

            string targetLibName = copyDialog.TargetListTitle;
            string sourceTitle = info.Title;
            string action = "Append"; // default — safe for both new and existing

            bool exists = await _spService.ListExistsAsync(targetUrl, targetLibName);
            var infoWin = new SPUtil.Views.OperationInfoWindow { Owner = Application.Current.MainWindow };

            if (!exists)
            {
                // New library — check for reserved names first
                if (ReservedLibNames.Contains(info.InternalName) || ReservedLibNames.Contains(sourceTitle))
                {
                    string errorMsg = $"The library '{sourceTitle}' is a system-reserved library (Internal Name: {info.InternalName}).\n\n" +
                                      "It cannot be created programmatically. Please ensure the required Features are activated " +
                                      "on the target site so SharePoint can create this library automatically.";

                    MessageBox.Show(errorMsg, "System Library Warning", MessageBoxButton.OK, MessageBoxImage.Stop);
                    ConnectionStatus = "Error: Reserved library name detected.";
                    return;
                }

                infoWin.Show();
                infoWin.UpdateMessage($"Creating Library structure: {targetLibName}...");
                bool created = await CreateListStructureAsync(info, targetUrl, targetLibName, templateId);
                infoWin.Close();
                if (!created) return;
            }
            else
            {
                // Library exists — ask user what to do
                var existsDialog = new SPUtil.Views.DocLibExistsActionDialog(targetLibName) { Owner = Application.Current.MainWindow };
                if (existsDialog.ShowDialog() != true) return;

                action = existsDialog.SelectedAction; // "Append", "Overwrite", or "Mirror"
                if (action == "Mirror")
                {
                    infoWin.Show();
                    infoWin.UpdateMessage($"Deleting library: {targetLibName}...");
                    await _spService.DeleteListAsync(targetUrl, targetLibName);

                    infoWin.UpdateMessage($"Creating Library structure: {targetLibName}...");
                    bool created = await CreateListStructureAsync(info, targetUrl, targetLibName, templateId);
                    infoWin.Close();
                    if (!created) return;
                }
            }

            // Step 2: Copy folder structure (always, regardless of action)
            var folderInfoWin = new SPUtil.Views.OperationInfoWindow { Owner = Application.Current.MainWindow };
            folderInfoWin.Show();
            var folderProgress = new Progress<string>(msg =>
                Application.Current.Dispatcher.Invoke(() => folderInfoWin.UpdateMessage(msg)));
            try
            {
                await _spService.CopyFolderStructureAsync(info.URL, targetUrl, sourceTitle, targetLibName, folderProgress);
            }
            finally
            {
                folderInfoWin.Close();
            }

            // Step 3: Permissions stub
            StatusMessage = "Setting up permissions (stub)...";
            // TODO: Call Change-PermissionsX here later

            // Step 4: Copy files (action determines behavior: Append / Overwrite / Mirror)
            if (withData)
            {
                // TODO: await ExecuteDocLibDataCopyAsync(info.URL, targetUrl, sourceTitle, targetLibName, action);
            }

            RightSiteNodes = await _spService.GetSiteStructureAsync(targetUrl);
            StatusMessage = "Finished copying library.";
        }
        private async Task<bool> CreateListStructureAsync(SPListInfo info, string targetUrl, string targetListName, int templateId)
        {
            try
            {
                
                 StatusMessage = "Анализ структуры и зависимостей...";
                 var fieldInfos = await _spService.GetFieldInfosFromSiteAsync(info.URL, info.Title);
                 var sourceViews = await _spService.GetListViewsAsync(info.URL, info.Title);
                
                //if (templateId == 100)
                //{
                    // Проверка Lookup
                    var missingLists = new List<string>();
                    foreach (var field in fieldInfos.Where(f => f.FieldType == "Lookup"))
                    {
                        bool targetExists = await _spService.ListExistsAsync(targetUrl, field.LookupListName);
                        if (!targetExists && !missingLists.Contains(field.LookupListName))
                            missingLists.Add(field.LookupListName);
                    }

                    if (missingLists.Count > 0)
                    {
                        string allMissing = string.Join("\n - ", missingLists);
                        System.Windows.MessageBox.Show($"Необходимы списки-зависимости:\n - {allMissing}", "Ошибка зависимостей");
                        return false;
                    }
                //}

                StatusMessage = "Создание списка на целевом сайте...";
                await _spService.CreateListFromSchemaAsync(targetUrl, info.InternalName, targetListName, fieldInfos, sourceViews, templateId);
                return true;
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка при создании структуры: {ex.Message}");
                return false;
            }
        }
    }
}
