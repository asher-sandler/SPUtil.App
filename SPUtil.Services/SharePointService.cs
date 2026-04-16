using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using Microsoft.SharePoint.DesignTime.Activities;
using Microsoft.SharePoint.JSGrid;
using Microsoft.Win32;
using SPUtil.Infrastructure;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Net;
using System.Security;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
//using SPUtil.Infrastructure;

namespace SPUtil.Services
{
	public partial class SharePointService : ISharePointService
	{
		/*
		public string GetConnectionStatus()
		{
            return "Connected";
		}
		*/
        private readonly SharePointCloneService _cloneService;

		public SharePointService(SharePointCloneService cloneService)
		{
			_cloneService = cloneService;
		}
		public string GetCurrentUsername()
		{
			try
			{
				string regPath = @"SOFTWARE\Microsoft\CrSiteAutomate";
				using (var key = Registry.CurrentUser.OpenSubKey(regPath))
				{
					// Пытаемся достать Param1, если ключа или значения нет — возвращаем "Unknown" или дефолт
					return key?.GetValue("Param1")?.ToString() ?? "Unknown";
				}
			}
			catch
			{
				return "Unknown";
			}
		}
		/*
		public string NormalizeUrl(string url)
		{
			if (string.IsNullOrWhiteSpace(url)) return url;

			try
			{
				Uri uri = new Uri(url);
				string host = uri.Host; // Например, portals2.ekmd.huji.ac.il
				string[] parts = host.Split('.');

				if (parts.Length > 0)
				{
					string firstPart = parts[0]; // portals2

					// Если первая часть заканчивается на '2', убираем её
					if (firstPart.EndsWith("2"))
					{
						parts[0] = firstPart.Remove(firstPart.Length - 1);

						// Собираем URL обратно
						var builder = new UriBuilder(uri);
						builder.Host = string.Join(".", parts);
						return builder.Uri.ToString().TrimEnd('/');
					}
				}
			}
			catch
			{
				// Если это не валидный URL, возвращаем как есть
			}

			return url.Trim();
		}
		*/
		
		
		private NetworkCredential GetCredentials()
		{
			/*
			string regPath = @"SOFTWARE\Microsoft\CrSiteAutomate";
			using (var key = Registry.CurrentUser.OpenSubKey(regPath))
			{
				string userName = key?.GetValue("Param1")?.ToString() ?? "Unknown";
				string encryptedHex = key?.GetValue("Param")?.ToString() ?? "";
				return new NetworkCredential(userName, DecryptFromPowerShell(encryptedHex), "ekmd");
			}
			*/
			return SPUtil.Infrastructure.SPUsingUtils.GetCredentials();
			
		}



        public async Task<AuthResult> ValidateConnectionAsync(string siteUrl)
        {
            try
            {
                using (var context = await GetContextAsync(siteUrl))
                {
                    context.Load(context.Web, w => w.Title);
                    await Task.Run(() => context.ExecuteQuery()); // Здесь происходит сетевой запрос
                    return AuthResult.Success;
                }
            }
            catch (Exception ex)
            {
                // Проверяем само исключение И внутреннее исключение (на всякий случай)
                var webEx = (ex as System.Net.WebException) ?? (ex.InnerException as System.Net.WebException);

                if (webEx != null && webEx.Response is System.Net.HttpWebResponse response)
                {
                    int code = (int)response.StatusCode;
                    System.Diagnostics.Debug.WriteLine($"[SP_AUTH_DEBUG] Site: {siteUrl} | HTTP Error: {code}");

                    switch (response.StatusCode)
                    {
                        case System.Net.HttpStatusCode.Unauthorized: // 401
                            return AuthResult.InvalidCredentials;
                        case System.Net.HttpStatusCode.Forbidden:    // 403
                            return AuthResult.AccessDenied;
                        case System.Net.HttpStatusCode.NotFound:     // 404
                            return AuthResult.SiteNotFound;
                    }
                }

                // Если это специфическая ошибка SharePoint ServerException
                if (ex is Microsoft.SharePoint.Client.ServerException serverEx)
                {
                    System.Diagnostics.Debug.WriteLine($"[SP_AUTH_DEBUG] ServerException: {serverEx.Message}");
                    if (serverEx.Message.Contains("Access denied") || serverEx.Message.Contains("У вас нет прав"))
                        return AuthResult.AccessDenied;
                }

                System.Diagnostics.Debug.WriteLine($"[SP_AUTH_DEBUG] General Error: {ex.Message}");
                return AuthResult.Error;
            }
        }
        // Вспомогательный метод для создания контекста (решает ошибку CS0103)
        private async Task<ClientContext> GetContextAsync(string siteUrl)
		{
			return await Task.Run(() =>
			{
				var context = new  ClientContext(SPUsingUtils.NormalizeUrl(siteUrl));
				context.Credentials = GetCredentials();
				return context;
			});
		}

		public async Task<ObservableCollection<SPNode>> GetSiteStructureAsync(string siteUrl)
		{
			return await Task.Run(async () =>
			{
				var nodes = new ObservableCollection<SPNode>();
				try
				{
					using (var context = await GetContextAsync(siteUrl))
					{
						Web web = context.Web;
						context.Load(web, w => w.Title);
						context.Load(web.Lists, lists => lists.Include(
							l => l.Title,
							l => l.Id,
							l => l.BaseTemplate,
							l => l.Hidden));

						context.ExecuteQuery();
						//GetConnectionStatus();
						foreach (var list in web.Lists.Where(l => !l.Hidden))
						{
							nodes.Add(new SPNode
							{
								Title = list.Title,
								Type = SharePointObjectType.List,
								Path = list.Id.ToString(),
								Tag = list.BaseTemplate.ToString()
							});
						}
					}
				}
				catch { }
				return nodes;
			});
		}


		public async Task<List<SPFileData>> GetLibraryItemsAsync(string siteUrl, string listId)
		{
			return await Task.Run(async () =>
			{
				using var context = await GetContextAsync(siteUrl);
				Guid g = new Guid(listId);
				List list = context.Web.Lists.GetById(g);

				CamlQuery query = CamlQuery.CreateAllItemsQuery();
				ListItemCollection items = list.GetItems(query);

				context.Load(items, icol => icol.Include(
					i => i.FileSystemObjectType,
					i => i["FileLeafRef"],
					i => i["FileRef"],
					i => i["Modified"]));

				context.ExecuteQuery();

				return items.ToList().Select(item =>
				{
					var isFolder = item.FileSystemObjectType == Microsoft.SharePoint.Client.FileSystemObjectType.Folder;
					return new SPFileData
					{
						Name = item["FileLeafRef"]?.ToString() ?? "",
						FullPath = item["FileRef"]?.ToString() ?? "", // Используем FullPath
						IsFolder = isFolder,
						Modified = item["Modified"] != null ? (DateTime)item["Modified"] : DateTime.MinValue,
						Size = 0
					};
				}).ToList();
			});
		}
		public async Task<bool> ListExistsAsync(string siteUrl, string listTitle)
		{
			return await Task.Run(async () =>
			{
				try
				{
					using var context = await GetContextAsync(siteUrl);
					// Пытаемся получить список по имени (Title)
					var lists = context.Web.Lists;
					context.Load(lists, l => l.Include(list => list.Title));
					context.ExecuteQuery();

					return lists.AsEnumerable().Any(l =>
						l.Title.Equals(listTitle, StringComparison.OrdinalIgnoreCase));
				}
				catch
				{
					// Если сайт недоступен или ошибка — считаем, что создавать нельзя (или обрабатываем иначе)
					return true;
				}
			});
		}

		public async Task<string> GetDetailedInfoAsync(string siteUrl, string listId, int templateId)
		{
			try
			{
				using var context = await GetContextAsync(siteUrl);
				Guid g = new Guid(listId);
				List list = context.Web.Lists.GetById(g);
				context.Load(list, l => l.Title, l => l.ItemCount);
				context.ExecuteQuery();

				return $"Список: {list.Title}\nТип: {templateId}\nЭлементов: {list.ItemCount}";
			}
			catch (Exception ex) { return $"Ошибка: {ex.Message}"; }
		}
		/*
		private SecureString DecryptFromPowerShell(string hexString)
		{
			if (string.IsNullOrEmpty(hexString)) return new SecureString();
			byte[] encryptedBytes = Enumerable.Range(0, hexString.Length / 2)
				.Select(x => Convert.ToByte(hexString.Substring(x * 2, 2), 16)).ToArray();
			byte[] decryptedBytes = ProtectedData.Unprotect(encryptedBytes, null, DataProtectionScope.CurrentUser);
			string plainText = Encoding.Unicode.GetString(decryptedBytes);
			var secureString = new SecureString();
			foreach (char c in plainText) secureString.AppendChar(c);
			secureString.MakeReadOnly();
			return secureString;
		}
		*/
		public async Task<List<SPFileData>> GetPageItemsAsync(string siteUrl, string listId)
		{
			return await Task.Run(async () =>
			{
				using var context = await GetContextAsync(siteUrl);
				var list = context.Web.Lists.GetById(new Guid(listId));

				CamlQuery query = new CamlQuery();
				// Рекурсивный поиск файлов и папок
				query.ViewXml = @"<View Scope='RecursiveAll'><Query></Query></View>";

				ListItemCollection items = list.GetItems(query);
				context.Load(items, icol => icol.Include(
					i => i.FileSystemObjectType,
					i => i["FileLeafRef"],
					i => i["FileRef"],
					i => i["Modified"]));
				context.ExecuteQuery();

				return items.ToList().Select(item => new SPFileData
				{
					Name = item["FileLeafRef"]?.ToString() ?? "",
					FullPath = item["FileRef"]?.ToString() ?? "",
					IsFolder = item.FileSystemObjectType == FileSystemObjectType.Folder,
					Modified = item["Modified"] is DateTime dt ? dt : DateTime.MinValue
				}).ToList();
			});
		}

		public async Task<List<SPWebPartData>> GetWebPartsAsync(string siteUrl, string fileRelativeUrl)
		{
			return await Task.Run(async () =>
			{
				using var context = await GetContextAsync(siteUrl);
				// Получаем файл и его WebPartManager
				var file = context.Web.GetFileByServerRelativeUrl(fileRelativeUrl);
				var wpm = file.GetLimitedWebPartManager(PersonalizationScope.Shared);

				context.Load(wpm.WebParts, wps => wps.Include(
					wp => wp.WebPart.Title,
					wp => wp.WebPart.Properties,
					wp => wp.Id));
				context.ExecuteQuery();

				var result = new List<SPWebPartData>();
				foreach (var wpDefinition in wpm.WebParts)
				{
					var wp = new SPWebPartData
					{
						Title = wpDefinition.WebPart.Title,
						Id = wpDefinition.Id.ToString(),
						Properties = wpDefinition.WebPart.Properties.FieldValues
									 .ToDictionary(k => k.Key, v => v.Value?.ToString() ?? "null")
					};
					result.Add(wp);
				}
				return result;
			});
		}
/*
		 public async Task<List<SPListItemData>> GetListItemsByTitleAsync(string siteUrl, string listTitle, IProgress<int> progress)
		{
			return await Task.Run(async () =>
			{
				using var ctx = await GetContextAsync(siteUrl);
				
					var list = ctx.Web.Lists.GetByTitle(listTitle);
					ctx.Load(list, l => l.ItemCount);
					ctx.ExecuteQuery();

					int total = list.ItemCount;
					int loaded = 0;
					var allItems = new List<SPListItemData>();

					// Вместо query.RowLimit = 500; делаем так:
					CamlQuery query = new CamlQuery();
					query.ViewXml = @"<View Scope='RecursiveAll'>
										<Query>
											<OrderBy><FieldRef Name='ID' Ascending='TRUE' /></OrderBy>
										</Query>
										<RowLimit>10</RowLimit>
									  </View>";
					do
					{
						var items = list.GetItems(query);
						ctx.Load(items);
						ctx.ExecuteQuery();

						query.ListItemCollectionPosition = items.ListItemCollectionPosition;
						
						foreach(var i in items) 
						{
							allItems.Add(new SPListItemData { Id = i.Id, Values = i.FieldValues });
						}

						loaded += items.Count;
						if (total > 0)
						{
							progress?.Report((loaded * 100) / total); // Отправляем прогресс в UI
						}
						else
						{
							progress?.Report(100); // Если элементов 0, прогресс сразу 100%
						}

                } while (query.ListItemCollectionPosition != null);

					return allItems;
				
			});
		} 
*/
		public async Task<List<SPListItemData>> GetListItemsByTitleAsync(
			string siteUrl, 
			string listTitle, 
			IProgress<int> progress, 
			CancellationToken ct)
		{
			return await Task.Run(async () =>
			{
				using var ctx = await GetContextAsync(siteUrl);
				var list = ctx.Web.Lists.GetByTitle(listTitle);
				ctx.Load(list, l => l.ItemCount);
				ctx.ExecuteQuery(); // В вашей версии только синхронный

				int total = list.ItemCount;
				int loaded = 0;
				var allItems = new List<SPListItemData>();

				CamlQuery query = new CamlQuery();
				query.ViewXml = @"<View Scope='RecursiveAll'>
									<Query>
										<OrderBy><FieldRef Name='ID' Ascending='TRUE' /></OrderBy>
									</Query>
									<RowLimit>19</RowLimit> 
								  </View>";
				do
				{
					ct.ThrowIfCancellationRequested();

					var items = list.GetItems(query);
					ctx.Load(items);
					ctx.ExecuteQuery(); // Блокирующий вызов

					query.ListItemCollectionPosition = items.ListItemCollectionPosition;
					
					foreach(var i in items) 
					{
						ct.ThrowIfCancellationRequested();
						allItems.Add(new SPListItemData { Id = i.Id, Values = i.FieldValues });
					}

					loaded += items.Count;
					if (total > 0)
					{
						progress?.Report((loaded * 100) / total);
					}

					// Даем WPF время обновить ProgressBar перед следующим тяжелым ExecuteQuery
					await Task.Yield(); 

				} while (query.ListItemCollectionPosition != null);

				return allItems;
			}, ct);
		}

       public async Task<List<SPListItemData>> GetListItemsByIDAsync(string siteUrl, string listId)
		{
			return await Task.Run(async () =>
			{
				using var context = await GetContextAsync(siteUrl);
				var list = context.Web.Lists.GetById(new Guid(listId));

				// Базовый запрос для Default View
				var query = CamlQuery.CreateAllItemsQuery();
				var items = list.GetItems(query);

				// Загружаем только необходимые свойства без null-propagating операторов
				context.Load(items, icol => icol.Include(
					i => i.Id,
					i => i["Title"]));

				// ИСПОЛЬЗУЕМ СИНХРОННЫЙ МЕТОД
				context.ExecuteQuery();

				return items.AsEnumerable().Select(i => new SPListItemData
				{
					Id = i.Id,
					Title = i["Title"] != null ? i["Title"].ToString() : "Без названия"
				}).ToList();
			});
		}
		//
		public async Task<string> GetListNameByIdAsync(string siteUrl, string listId)
		{
			if (string.IsNullOrEmpty(listId)) return "Unknown";

			try
			{
				// Убираем фигурные скобки и парсим строку в Guid
				Guid g = new Guid(listId.Trim('{', '}'));
				
				// Вызываем основной типизированный метод
				return await GetListTitleByGuidAsync(siteUrl, g);
			}
			catch
			{
				return $"List not found ({listId})";
			}
		}

		public async Task<string> GetListTitleByGuidAsync(string siteUrl, Guid listGuid)
		{
			// Используем GetContextAsync (который уже асинхронный внутри вашего сервиса)
			using (var ctx = await GetContextAsync(siteUrl))
			{
				var list = ctx.Web.Lists.GetById(listGuid);
				ctx.Load(list, l => l.Title);
				
				// Выполняем запрос в фоновом потоке, чтобы не блокировать UI
				await Task.Run(() => ctx.ExecuteQuery());
				
				return list.Title;
			}
		}
		public async Task<List<string>> GetListSchemaAsync(string siteUrl, string listTitle)
		{
			try 
			{
                // Вызываем новый метод, который берет SchemaXml напрямую из SharePoint
                // без фильтрации и преобразований
                //return await _cloneService.GetAllRawFieldSchemasAsync(siteUrl, listTitle);
                return await GetAllRawFieldSchemasAsync(siteUrl, listTitle);
            }
			catch (Exception ex)
			{
				// Логирование ошибки, если нужно
				throw new Exception($"Не удалось получить полную схему списка: {ex.Message}");
			}
		}
        private async Task<List<string>> GetAllRawFieldSchemasAsync(string siteUrl, string listTitle)
        {
            var schemas = new List<string>();
            using (var ctx = await GetContextAsync(siteUrl))
            {
                // Здесь должна быть ваша настройка Credentials

                Web web = ctx.Web;
                List list = web.Lists.GetByTitle(listTitle);
                // Загружаем расширенный набор свойств, включая ReadOnlyField и Formula
                ctx.Load(list.Fields, fs => fs.Include(
                    f => f.Title,
                    f => f.InternalName,
                    f => f.StaticName,
                    f => f.FieldTypeKind,
                    f => f.Required,
                    f => f.ReadOnlyField, // Нужно для фильтрации системных полей
                    f => f.Hidden,        // Нужно для фильтрации скрытых полей
                    f => f.SchemaXml      // Резервный источник данных
                ));


                //ctx.Load(list.Fields, fs => fs.Include(f => f.SchemaXml));
                await Task.Run(() => ctx.ExecuteQuery());


                foreach (var field in list.Fields)
                {
                    if (field.InternalName.Equals("Id", StringComparison.OrdinalIgnoreCase))
                    {
                        schemas.Add(field.SchemaXml);

                    }

                }
                foreach (var field in list.Fields)
                {
                    if (field.InternalName.Equals("Title", StringComparison.OrdinalIgnoreCase))
                    {
                        schemas.Add(field.SchemaXml);

                    }

                }

                foreach (var field in list.Fields)
                {
                    // Если поле только для чтения, мы его пропускаем, КРОМЕ вычисляемых полей
                    // Пропускаем системные ReadOnly поля, но ОСТАВЛЯЕМ Id (для маппинга) и Calculated (для формул)
                    if (field.InternalName.Equals("Id", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (field.InternalName.Equals("Title", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (field.InternalName.Equals("ContentType", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    if (field.InternalName.Equals("Attachments", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    // 1. Список имен, которые мы хотим обрабатывать (не пропускать), даже если они ReadOnly
                    // 1. Проверяем, является ли поле системным исключением (которые мы хотим оставить)
                    bool isSystemException =
                        field.InternalName.Equals("Created", StringComparison.OrdinalIgnoreCase) ||
                        field.InternalName.Equals("Author", StringComparison.OrdinalIgnoreCase) ||
                        field.InternalName.Equals("Modified", StringComparison.OrdinalIgnoreCase) ||
                        field.InternalName.Equals("Editor", StringComparison.OrdinalIgnoreCase);

                    // 2. Логика фильтрации:
                    // МЫ ПРОПУСКАЕМ (continue) поле, ТОЛЬКО если оно:
                    // - Помечено как ReadOnly
                    // - И при этом НЕ является вычисляемым (Calculated)
                    // - И при этом НЕ является системным исключением (Id, Created и т.д.)
                    if (field.ReadOnlyField && field.FieldTypeKind != FieldType.Calculated && !isSystemException)
                    {
                        continue;
                    }

                    // 3. Если мы дошли до сюда, поле проходит фильтр и добавляется в список
                    // (Здесь должен идти ваш Switch по типам полей и в конце fieldInfos.Add(info))  
                    // Скрытые поля и системные префиксы игнорируем всегда
                    if (!field.InternalName.StartsWith("_x")) // Hebrew Names
                    {
                        if (field.Hidden || field.InternalName.StartsWith("_") || field.InternalName.StartsWith("vti_"))
                        {
                            continue;
                        }
                    }

                    if (!string.IsNullOrEmpty(field.SchemaXml))
                    {
                        schemas.Add(field.SchemaXml);
                    }
                }



            }
            return schemas;
        }
        public async Task CreateDocLibAsync(string siteUrl, string listName, string displayName = "")
        {
            // Используем ваш вспомогательный метод для получения контекста
            using (var context = await GetContextAsync(siteUrl))
            {
                Web web = context.Web;
                context.Load(web, w => w.Lists);
                context.ExecuteQuery();

                // Проверка на существование (аналог Check-ListExists в PS)
                // Ищем либо по системному имени, либо по отображаемому
                string searchTitle = string.IsNullOrEmpty(displayName) ? listName : displayName;
                bool exists = false;

                try
                {
                    List existingList = web.Lists.GetByTitle(searchTitle);
                    context.Load(existingList);
                    context.ExecuteQuery();
                    exists = true;
                }
                catch { /* Список не найден */ }

                if (!exists)
                {
                    // Настройка информации о создании (ListCreationInformation)
                    ListCreationInformation listInfo = new ListCreationInformation
                    {
                        Title = listName,
                        TemplateType = 101 // Document Library
                    };

                    List newList = web.Lists.Add(listInfo);

                    // Если задано отображаемое имя (различается от системного)
                    if (!string.IsNullOrEmpty(displayName))
                    {
                        newList.Title = displayName;
                    }

                    // Настройки из вашего скрипта
                    newList.OnQuickLaunch = false; // Убрать из бокового меню

                    // Если нужно задать современный интерфейс (закомментировано в PS, но полезно)
                    // newList.ListExperienceOptions = ListExperience.NewExperience;

                    newList.Update();
                    context.ExecuteQuery();

                    System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] DocLib '{searchTitle}' created successfully.");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] DocLib '{searchTitle}' already exists. Skipping creation.");
                }
            }
        }
        public async Task CreateListFromSchemaAsync(string targetUrl, string internalListName, string newlistTitle, List<FieldInfo> sourceFields, List<SPViewData> sourceViews, int listType = 100)
        {
            await Task.Run(async () =>
            {
                using (var ctx = await GetContextAsync(targetUrl))
                {
                    // 1. Создаем список
                    var newList = await CreateListAsync(ctx, internalListName, newlistTitle, listType);

                    // 2. Инициализируем существующие поля и представления
                    ctx.Load(newList.Fields, fs => fs.Include(f => f.InternalName, f => f.Id));
                    ctx.Load(newList.Views, vs => vs.Include(v => v.Title));
                    ctx.ExecuteQuery();

                    var existingFields = newList.Fields.Select(f => f.InternalName).ToList();

                    // 3. ОТДЕЛЬНЫЙ МЕТОД: Обработка Lookup и Dependency Lookup
                    await ProcessAndCreateLookupFieldsAsync(ctx, newList, sourceFields, existingFields);

                    // 4. ДОБАВЛЯЕМ ОСТАЛЬНЫЕ ОБЫЧНЫЕ ПОЛЯ (кроме Calculated и Lookup)
                    var otherNormalFields = sourceFields.Where(f =>
                        !f.FieldType.Equals("Calculated", StringComparison.OrdinalIgnoreCase) &&
                        !f.FieldType.Equals("Lookup", StringComparison.OrdinalIgnoreCase) &&
                        !f.FieldType.Equals("LookupMulti", StringComparison.OrdinalIgnoreCase)).ToList();

                    foreach (var field in otherNormalFields)
                    {
                        if (existingFields.Contains(field.Name, StringComparer.OrdinalIgnoreCase)) continue;

                        string fieldXml = _cloneService.GenerateFieldXml(field);

                        // Спец-обработка TargetTo (как в вашем исходном коде)
                        if (field.DisplayName == "Target Audiences" && field.FieldType == "Invalid")
                        {
                            fieldXml = @"<Field Type=""TargetTo"" DisplayName=""Target Audiences"" Required=""FALSE""><Customization><ArrayOfProperty><Property><Name>AllowGlobalAudience</Name><Value xmlns:q1=""http://www.w3.org/2001/XMLSchema"" p4:type=""q1:boolean"" xmlns:p4=""http://www.w3.org/2001/XMLSchema-instance"">true</Value></Property><Property><Name>AllowDL</Name><Value xmlns:q2=""http://www.w3.org/2001/XMLSchema"" p4:type=""q2:boolean"" xmlns:p4=""http://www.w3.org/2001/XMLSchema-instance"">true</Value></Property><Property><Name>AllowSPGroup</Name><Value xmlns:q3=""http://www.w3.org/2001/XMLSchema"" p4:type=""q3:boolean"" xmlns:p4=""http://www.w3.org/2001/XMLSchema-instance"">true</Value></Property></ArrayOfProperty></Customization></Field>";
                        }

                        newList.Fields.AddFieldAsXml(fieldXml, true, AddFieldOptions.AddFieldInternalNameHint);
                    }
                    newList.Update();
                    ctx.ExecuteQuery();

                    // 5. ДОБАВЛЯЕМ ВЫЧИСЛЯЕМЫЕ ПОЛЯ (Многопроходный цикл)
                    var calculatedFields = sourceFields.Where(f => f.FieldType.Equals("Calculated", StringComparison.OrdinalIgnoreCase)).ToList();
                    await CreateCalculatedFieldsAsync(ctx, newList, calculatedFields);

                    // 6. ДОБАВЛЯЕМ ИЛИ ОБНОВЛЯЕМ ПРЕДСТАВЛЕНИЯ
                    await CreateListViewsAsync(ctx, newList, sourceViews);

                    newList.Update();
                    ctx.ExecuteQuery();
                }
            });
        }
        private async Task ProcessAndCreateLookupFieldsAsync(ClientContext ctx, List newList, List<FieldInfo> sourceFields, List<string> existingFields)
        {
            var lookups = sourceFields.Where(f => f.FieldType.StartsWith("Lookup", StringComparison.OrdinalIgnoreCase)).ToList();
            var fieldGuidMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Этап А: Создаем основные Lookup-поля
            foreach (var field in lookups.Where(f => !f.IsDependentLookup))
            {
                if (existingFields.Contains(field.Name, StringComparer.OrdinalIgnoreCase)) continue;

                try
                {
                    // Маппинг ID списка
                    var targetLookupList = ctx.Web.Lists.GetByTitle(field.LookupListName);
                    ctx.Load(targetLookupList, l => l.Id, l => l.Title);
                    ctx.ExecuteQuery();
                    field.LookupListId = targetLookupList.Id.ToString();
                    field.LookupWebId = string.Empty;

                    string xml = _cloneService.GenerateFieldXml(field);
                    System.Diagnostics.Debug.WriteLine($"XML Lookup {xml} ");
                    var createdField = newList.Fields.AddFieldAsXml(xml, true, AddFieldOptions.AddFieldInternalNameHint);
                    ctx.Load(createdField, f => f.Id);
                    ctx.ExecuteQuery();


                    // Запоминаем ID созданного поля (используем ID источника как ключ)
                    // ЗАПОМИНАЕМ: Ключ = "ИмяСписка:ИмяПоля"
                    string key = $"{field.LookupListName}:{field.Name}";
                    fieldGuidMap[key] = createdField.Id.ToString();

                    System.Diagnostics.Debug.WriteLine($"[MAP] Сохранили ID для {key} -> {createdField.Id}");
                }

                catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"Lookup Error {field.Name}: {ex.Message}"); }
            }

            // Этап Б: Создаем зависимые (Dependency) Lookup-поля (например, sort1:Code)
            foreach (var field in lookups.Where(f => f.IsDependentLookup))
            {
                if (existingFields.Contains(field.Name, StringComparer.OrdinalIgnoreCase)) continue;

                string decodedName = field.Name.Replace("_x003a_", ":");

                string parentFieldName = string.Empty;
                string showField = string.Empty;

                
                //string lookupShowField = "Title"; // по умолчанию

				if (decodedName.Contains(':'))
				{
                    var parts = decodedName.Split(':');
                    parentFieldName = parts[0]; // "Country"
                    showField = parts[1];       // "Title"



                    // Теперь формируем ключ для поиска в словаре, который наполнили в Этапе А
                    string parentKey = $"{field.LookupListName}:{parentFieldName}";

					if (fieldGuidMap.TryGetValue(parentKey, out string newParentId))
					{
						try
						{
							// Нашли родителя по составному ключу!
							field.FieldRef = newParentId;

							// Получаем ID списка для атрибута List
							var targetLookupList = ctx.Web.Lists.GetByTitle(field.LookupListName);
							ctx.Load(targetLookupList, l => l.Id);
							ctx.ExecuteQuery();
							field.LookupListId = targetLookupList.Id.ToString();

							string xml = _cloneService.GenerateFieldXml(field);
                            System.Diagnostics.Debug.WriteLine($"XML Dependency Lookup {xml} ");

                            newList.Fields.AddFieldAsXml(xml, true, AddFieldOptions.AddFieldInternalNameHint);
						}
						catch (Exception ex)
						{
							System.Diagnostics.Debug.WriteLine($"Dep-Lookup Error {field.Name}: {ex.Message}");
						}
					}
				}
            }

            newList.Update();
            ctx.ExecuteQuery();
        }
        private async Task CreateListViewsAsync(ClientContext ctx, Microsoft.SharePoint.Client.List newList, List<SPViewData> sourceViews)
        {
            foreach (var viewData in sourceViews)
            {
                try
                {
                    Microsoft.SharePoint.Client.View targetView;
                    // Проверяем, существует ли уже вьюха
                    var existingViews = ctx.LoadQuery(newList.Views.Where(v => v.Title == viewData.Title));
                    ctx.ExecuteQuery();
                    var existingView = existingViews.FirstOrDefault();

                    if (existingView != null)
                    {
                        targetView = existingView;
                    }
                    else
                    {
                        ViewCreationInformation vInfo = new ViewCreationInformation
                        {
                            Title = viewData.Title,
                            PersonalView = false
                        };
                        targetView = newList.Views.Add(vInfo);
                    }

                    targetView.ViewQuery = viewData.ViewQuery;
                    targetView.DefaultView = viewData.DefaultView;

                    if (!string.IsNullOrEmpty(viewData.Aggregations))
                    {
                        targetView.Aggregations = viewData.Aggregations;
                    }

                    // Настройка полей во вьюхе
                    ctx.Load(targetView.ViewFields);
                    ctx.ExecuteQuery();

                    // Очищаем старые поля (кроме первого, чтобы не упало)
                    string firstField = viewData.ViewFields?.FirstOrDefault() ?? "LinkTitle";
                    var currentFields = targetView.ViewFields.ToArray();
                    foreach (var fName in currentFields)
                    {
                        if (!fName.Equals(firstField, StringComparison.OrdinalIgnoreCase))
                        {
                            targetView.ViewFields.Remove(fName);
                        }
                    }
                    targetView.Update();

                    // Добавляем нужные поля
                    if (viewData.ViewFields != null && viewData.ViewFields.Length > 0)
                    {
                        foreach (var fName in viewData.ViewFields)
                        {
                            if (fName.Equals(firstField, StringComparison.OrdinalIgnoreCase)) continue;
                            targetView.ViewFields.Add(fName);
                        }
                    }

                    targetView.Update();
                    ctx.ExecuteQuery();
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] Error creating view '{viewData.Title}': {ex.Message}");
                }
            }
        }
        private async Task CreateCalculatedFieldsAsync(ClientContext ctx, Microsoft.SharePoint.Client.List newList, List<FieldInfo> calculatedFields)
        {
            int maxAttempts = 5;
            int attempt = 0;

            while (calculatedFields.Count > 0 && attempt < maxAttempts)
            {
                var succeeded = new List<FieldInfo>();
                foreach (var calcField in calculatedFields)
                {
                    try
                    {
                        // Генерируем XML для вычисляемого поля
                        string calcXML = _cloneService.GenerateFieldXml(calcField);

                        System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] Attempting Calc Field: {calcField.DisplayName}");

                        newList.Fields.AddFieldAsXml(calcXML, true, AddFieldOptions.DefaultValue);
                        newList.Update();
                        ctx.ExecuteQuery();

                        succeeded.Add(calcField);
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] Attempt {attempt} failed for {calcField.DisplayName}: {ex.Message}");
                    }
                }

                // Удаляем те, что удалось создать
                foreach (var item in succeeded)
                {
                    calculatedFields.Remove(item);
                }
                attempt++;
            }
        }
        public async Task<Guid> GetListIdByTitleAsync(string siteUrl, string listTitle)
		{
			using (var ctx = await GetContextAsync(siteUrl))
			{
				var list = ctx.Web.Lists.GetByTitle(listTitle);
				ctx.Load(list, l => l.Id);
				await Task.Run(() => ctx.ExecuteQuery());
				return list.Id;
			}
		}		
        /*
        public async Task CreateListFromSchemaAsync(string targetUrl, string internalListName, string newlistTitle, List<FieldInfo> sourceFields, List<SPViewData> sourceViews,int listType=100)
		{
			await Task.Run(async () =>
			{
				using (var ctx = await GetContextAsync(targetUrl))
				{
                    System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] List Internal name: '{internalListName}' .");

                    // 1. Создаем список
                    var newList = await CreateListAsync(ctx, internalListName, newlistTitle, listType);

					// 2. Получаем существующие поля, чтобы избежать конфликтов (Title, ID и т.д.)
					ctx.Load(newList.Fields, fs => fs.Include(f => f.InternalName));
					ctx.Load(newList.Views, vs => vs.Include(v => v.Title));
					ctx.ExecuteQuery();

					var existingFields = newList.Fields.Select(f => f.InternalName).ToList();
					// --- 3. ОБРАБОТКА LOOKUP ПОЛЕЙ (ПОДМЕНА ID) ---
					foreach (var field in sourceFields.Where(f => f.FieldType.Equals("Lookup", StringComparison.OrdinalIgnoreCase)))
					{
						try
						{
							// Ищем GUID аналогичного списка на целевом сайте по имени
							var targetLookupList = ctx.Web.Lists.GetByTitle(field.LookupListName);
							ctx.Load(targetLookupList, l => l.Id);
							ctx.ExecuteQuery();

							// Обновляем ID в данных поля перед генерацией XML
							field.LookupListId = targetLookupList.Id.ToString();
							field.LookupWebId = string.Empty; // Сбрасываем, чтобы поиск шел внутри текущего сайта
						}
						catch (Exception ex)
						{
							System.Diagnostics.Debug.WriteLine($"Ошибка маппинга Lookup для {field.Name}: {ex.Message}");
						}
					}
					// --- НОВОЕ: Разделяем поля на обычные и вычисляемые ---
					var normalFields = sourceFields.Where(f => !f.FieldType.Equals("Calculated", StringComparison.OrdinalIgnoreCase)).ToList();
					var calculatedFields = sourceFields.Where(f => f.FieldType.Equals("Calculated", StringComparison.OrdinalIgnoreCase)).ToList();

					// 3. ДОБАВЛЯЕМ ОБЫЧНЫЕ ПОЛЯ
					foreach (var field in normalFields)
					{
						if (existingFields.Contains(field.Name, StringComparer.OrdinalIgnoreCase)) continue;
						
                        string fieldXml = _cloneService.GenerateFieldXml(field);
                        if (field.DisplayName == "Target Audiences" && field.FieldType == "Invalid")
                        {
                            System.Diagnostics.Debug.WriteLine($"Field Target: {field.DisplayName}");
                            fieldXml = @"
<Field Type=""TargetTo"" 
       DisplayName=""Target Audiences"" 
       Required=""FALSE"" >
  <Customization>
    <ArrayOfProperty>
      <Property>
        <Name>AllowGlobalAudience</Name>
        <Value xmlns:q1=""http://www.w3.org/2001/XMLSchema"" p4:type=""q1:boolean"" xmlns:p4=""http://www.w3.org/2001/XMLSchema-instance"">true</Value>
      </Property>
      <Property>
        <Name>AllowDL</Name>
        <Value xmlns:q2=""http://www.w3.org/2001/XMLSchema"" p4:type=""q2:boolean"" xmlns:p4=""http://www.w3.org/2001/XMLSchema-instance"">true</Value>
      </Property>
      <Property>
        <Name>AllowSPGroup</Name>
        <Value xmlns:q3=""http://www.w3.org/2001/XMLSchema"" p4:type=""q3:boolean"" xmlns:p4=""http://www.w3.org/2001/XMLSchema-instance"">true</Value>
      </Property>
    </ArrayOfProperty>
  </Customization>
</Field>";
                        }

                        System.Diagnostics.Debug.WriteLine($"Field XML: {fieldXml}");
							
						newList.Fields.AddFieldAsXml(fieldXml, true, Microsoft.SharePoint.Client.AddFieldOptions.AddFieldInternalNameHint);
					}
					newList.Update();
					ctx.ExecuteQuery();

					// --- НОВОЕ: ДОБАВЛЯЕМ ВЫЧИСЛЯЕМЫЕ ПОЛЯ (Многопроходный цикл) ---
					int maxAttempts = 5; 
					int attempt = 0;

					while (calculatedFields.Count > 0 && attempt < maxAttempts)
					{
						var succeeded = new List<FieldInfo>();
						foreach (var calcField in calculatedFields)
						{
							try
							{
								// Используем SchemaXml, так как в нем хранится готовая формула и ResultType
								// Если SchemaXml пуст, используем генератор
								string calcXML = _cloneService.GenerateFieldXml(calcField);
                                //calcXML = calcField.BuildXml();
                                System.Diagnostics.Debug.WriteLine($"Field XML (calc): {calcXML}");

                                //string fXml = !string.IsNullOrEmpty(calcField.SchemaXml) 
								//			  ? calcField.SchemaXml 
								//			  : _cloneService.GenerateFieldXml(calcField);
								//System.Diagnostics.Debug.WriteLine($"Field XML (calc): {fXml}");
											  

								newList.Fields.AddFieldAsXml(calcXML, true, Microsoft.SharePoint.Client.AddFieldOptions.DefaultValue);
								newList.Update();
								ctx.ExecuteQuery(); 
								
								succeeded.Add(calcField);
							}
							catch (Exception ex)
							{
								// Если ошибка — возможно, поле-родитель еще не создано или ошибка в формуле
								System.Diagnostics.Debug.WriteLine($"Attempt {attempt} failed for {calcField.DisplayName}: {ex.Message}");
							}
						}

						// Удаляем успешно созданные из списка ожидания
						foreach (var item in succeeded)
						{
							calculatedFields.Remove(item);
						}
						attempt++;
					}

					// 4. ДОБАВЛЯЕМ ИЛИ ОБНОВЛЯЕМ ПРЕДСТАВЛЕНИЯ
					foreach (var viewData in sourceViews)
					{
						Microsoft.SharePoint.Client.View targetView;
						var existingView = newList.Views.Where(v => v.Title.Equals(viewData.Title, StringComparison.OrdinalIgnoreCase)).FirstOrDefault();

						if (existingView != null)
						{
							targetView = existingView;
						}
						else
						{
							Microsoft.SharePoint.Client.ViewCreationInformation vInfo = new Microsoft.SharePoint.Client.ViewCreationInformation
							{
								Title = viewData.Title,
								PersonalView = false
							};
							targetView = newList.Views.Add(vInfo);
						}

						targetView.ViewQuery = viewData.ViewQuery;
						targetView.DefaultView = viewData.DefaultView;

						if (!string.IsNullOrEmpty(viewData.Aggregations))
						{
							targetView.Aggregations = viewData.Aggregations;
						}

						ctx.Load(targetView.ViewFields);
						ctx.ExecuteQuery();

						string firstField = viewData.ViewFields?.FirstOrDefault() ?? "LinkTitle";
						var currentFields = targetView.ViewFields.ToArray();

						foreach (var fName in currentFields)
						{
							if (!fName.Equals(firstField, StringComparison.OrdinalIgnoreCase))
							{
								targetView.ViewFields.Remove(fName);
							}
						}
						targetView.Update();

						if (viewData.ViewFields != null && viewData.ViewFields.Length > 0)
						{
							if (!currentFields.Contains(firstField, StringComparer.OrdinalIgnoreCase))
							{
								targetView.ViewFields.Add(firstField);
							}

							for (int i = 1; i < viewData.ViewFields.Length; i++)
							{
								targetView.ViewFields.Add(viewData.ViewFields[i]);
							}
						}

						targetView.Update();
					}

					newList.Update();
					ctx.ExecuteQuery();
				}
			});
		}
		*/
        private async Task<List> CreateListAsync(ClientContext ctx, string internalName, string newlistTitle, int templateType)
        {
            ListCreationInformation creationInfo = new ListCreationInformation
            {
                Title = internalName,
                TemplateType = templateType
            };

            List newList = ctx.Web.Lists.Add(creationInfo);

            // --- ШАГ 2: Сразу меняем отображаемый заголовок на правильный ---
            // Это не изменит URL (RootFolder), но в интерфейсе будет красиво
            if (internalName != newlistTitle)
            {
                newList.Title = newlistTitle;
                newList.Update();
            }

            // Загружаем поля и виды для дальнейшей работы
            ctx.Load(newList.Fields, fs => fs.Include(f => f.InternalName));
            ctx.Load(newList.Views, vs => vs.Include(v => v.Title));

            // Выполняем создание и переименование одним запросом
            await Task.Run(() => ctx.ExecuteQuery());

            System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] List created. URL Name: '{internalName}', Display Title: '{newlistTitle}'");
            
            return newList;
        }

		public async Task ClearListItemsAsync(string siteUrl, string listTitle)
		{
			await Task.Run(async () =>
			{
				using (var ctx = await GetContextAsync(siteUrl))
				{
					List list = ctx.Web.Lists.GetByTitle(listTitle);
					
					// Используем простой запрос для удаления первых 500 найденных элементов
					CamlQuery query = new CamlQuery();
					query.ViewXml = @"<View Scope='RecursiveAll'>
										<Query><OrderBy><FieldRef Name='ID' Ascending='TRUE' /></OrderBy></Query>
										<RowLimit>500</RowLimit>
									  </View>";

					bool hasItems = true;
					while (hasItems)
					{
						ListItemCollection items = list.GetItems(query);
						ctx.Load(items, icol => icol.Include(i => i.Id));
						ctx.ExecuteQuery();

						if (items.Count > 0)
						{
							System.Diagnostics.Debug.WriteLine($"[CLEANUP] Удаление пакета: {items.Count} элементов...");
							for (int i = items.Count - 1; i >= 0; i--)
							{
								items[i].DeleteObject();
							}
							ctx.ExecuteQuery(); // Удаляем пакет
						}
						else
						{
							hasItems = false;
						}
					}
				}
			});
		}
        public async Task<SPListInfo> GetListDetailedInfoAsync(string siteUrl, string listTitle)
        {
            using (var ctx = await GetContextAsync(siteUrl))
            {
                List list = ctx.Web.Lists.GetByTitle(listTitle);

                ctx.Load(list,
                    l => l.Title,
                    l => l.EntityTypeName,
					l => l.ParentWebUrl,
                    l => l.BaseType,
                    l => l.BaseTemplate, // Добавили шаблон
                    l => l.ItemCount,
                    l => l.Created);
				ctx.Load(list.RootFolder, r => r.ServerRelativeUrl,r => r.Name);

                ctx.ExecuteQuery();

                bool isDocLib = list.BaseType == Microsoft.SharePoint.Client.BaseType.DocumentLibrary
                                || list.BaseTemplate == 101;

                return new SPListInfo
                {
					URL = siteUrl,
                    Title = list.Title,
                    InternalName = list.RootFolder.Name,
                    Type = isDocLib ? "Document Library" : "List",
                    ItemCount = list.ItemCount,
					ServerRelativeUrl = list.RootFolder.ServerRelativeUrl,
					ParentWebUrl = list.ParentWebUrl,
                    Created = list.Created,
                    BaseTemplate = list.BaseTemplate
                };
            }
        }
        public async Task DeleteListAsync(string siteUrl, string listTitle)
		{
			using (var ctx = await GetContextAsync(siteUrl))
			{
				List list = ctx.Web.Lists.GetByTitle(listTitle);
				list.DeleteObject();
				await Task.Run(() => ctx.ExecuteQuery());
			}
		}
        /// <summary>
        /// Copies the entire empty folder structure (any depth) from source doc library to target.
        /// Files are NOT copied — only folder hierarchy is recreated.
        /// </summary>
        public async Task CopyFolderStructureAsync(
            string sourceUrl,
            string targetUrl,
            string sourceLibraryTitle,
            string targetLibraryTitle,
            IProgress<string> progress = null)
        {
            await Task.Run(async () =>
            {
                using var sourceCtx = await GetContextAsync(sourceUrl);
                using var targetCtx = await GetContextAsync(targetUrl);

                // --- Step 1: Load source library root folder server-relative URL ---
                var sourceList = sourceCtx.Web.Lists.GetByTitle(sourceLibraryTitle);
                sourceCtx.Load(sourceList.RootFolder, r => r.ServerRelativeUrl);
                await Task.Run(() => sourceCtx.ExecuteQuery());

                string sourceRootUrl = sourceList.RootFolder.ServerRelativeUrl; // e.g. /sites/hr/Shared Documents

                // --- Step 2: Load target library root folder server-relative URL ---
                var targetList = targetCtx.Web.Lists.GetByTitle(targetLibraryTitle);
                targetCtx.Load(targetList.RootFolder, r => r.ServerRelativeUrl);
                await Task.Run(() => targetCtx.ExecuteQuery());

                string targetRootUrl = targetList.RootFolder.ServerRelativeUrl; // e.g. /sites/hr2/Shared Documents

                // --- Step 3: Fetch ALL folders from source (recursive, sorted by depth) ---
                var caml = new CamlQuery
                {
                    ViewXml = @"<View Scope='RecursiveAll'>
									<Query>
										<Where>
											<Eq>
												<FieldRef Name='FSObjType'/>
												<Value Type='Integer'>1</Value>
											</Eq>
										</Where>
										<OrderBy>
											<FieldRef Name='FileRef' Ascending='TRUE'/>
										</OrderBy>
									</Query>
								</View>"
                };

                var folderItems = sourceList.GetItems(caml);
                sourceCtx.Load(folderItems, items => items.Include(
                    i => i["FileRef"],      // full server-relative path
                    i => i["FileLeafRef"]   // folder name only
                ));
                await Task.Run(() => sourceCtx.ExecuteQuery());

                // Collect folder paths and sort by depth (parent before child)
                // Note: avoid ?. inside expression trees — use explicit null check instead
                var folderPaths = folderItems
                    .Cast<ListItem>()
                    .Select(i => i["FileRef"] != null ? i["FileRef"].ToString() : "")
                    .Where(p => !string.IsNullOrEmpty(p))
                    .OrderBy(p => p.Count(c => c == '/')) // shallow folders first
                    .ToList();

                System.Diagnostics.Debug.WriteLine($"[FOLDER_COPY] Found {folderPaths.Count} folders in '{sourceLibraryTitle}'");

                // --- Step 4: Recreate each folder on target ---
                foreach (var sourceFolderPath in folderPaths)
                {
                    // Build target path by replacing source root prefix with target root
                    if (!sourceFolderPath.StartsWith(sourceRootUrl, StringComparison.OrdinalIgnoreCase))
                    {
                        System.Diagnostics.Debug.WriteLine($"[FOLDER_COPY] Skipping unexpected path: {sourceFolderPath}");
                        continue;
                    }

                    // e.g. sourceFolderPath = /sites/hr/Shared Documents/2024/Q1
                    // relativePart         = /2024/Q1
                    // targetFolderPath     = /sites/hr2/Shared Documents/2024/Q1
                    string relativePart = sourceFolderPath.Substring(sourceRootUrl.Length);
                    string targetFolderPath = targetRootUrl + relativePart;

                    progress?.Report($"Creating folder: {relativePart}");
                    System.Diagnostics.Debug.WriteLine($"[FOLDER_COPY] Ensuring folder: {targetFolderPath}");

                    try
                    {
                        // Get the parent folder path and new folder name
                        // e.g. targetFolderPath = /sites/hr2/Shared Documents/2024/Q1
                        // parentPath            = /sites/hr2/Shared Documents/2024
                        // folderName            = Q1
                        string folderName = targetFolderPath.Substring(targetFolderPath.LastIndexOf('/') + 1);
                        string parentPath = targetFolderPath.Substring(0, targetFolderPath.LastIndexOf('/'));

                        // Folder.Exists is not available in this CSOM version —
                        // use try/catch: if parent not found SharePoint throws ServerException
                        Folder parentFolder = targetCtx.Web.GetFolderByServerRelativeUrl(parentPath);
                        targetCtx.Load(parentFolder, f => f.Name); // load any property to verify folder exists
                        await Task.Run(() => targetCtx.ExecuteQuery());

                        // If we reach here, parent folder exists — create the child
                        parentFolder.Folders.Add(folderName);
                        await Task.Run(() => targetCtx.ExecuteQuery());
                    }
                    catch (ServerException ex) when (ex.ServerErrorTypeName == "System.IO.FileNotFoundException")
                    {
                        // Parent folder doesn't exist — should not happen due to depth ordering, log and skip
                        System.Diagnostics.Debug.WriteLine($"[FOLDER_COPY] Parent not found, skipping: {targetFolderPath} | {ex.Message}");
                    }
                    catch (Exception ex)
                    {
                        // Folder may already exist or other non-critical error — log and continue
                        System.Diagnostics.Debug.WriteLine($"[FOLDER_COPY] Note for '{targetFolderPath}': {ex.Message}");
                    }
                }

                System.Diagnostics.Debug.WriteLine($"[FOLDER_COPY] Done. Processed {folderPaths.Count} folders.");
            });
        }
        public async Task CopyListItemsAsync(
			string sourceUrl, 
			string targetUrl, 
			string sourceTitle, 
			string targetListName, 
			string action,         
			IProgress<CopyProgressArgs> progress, 
			CancellationToken ct)
		{
			// --- ШАГ 1: ВЫЗОВ ОЧИСТКИ ПЕРЕД КОПИРОВАНИЕМ ---
			if (action == "Overwrite")
			{
				await ClearListItemsAsync(targetUrl, targetListName);
			}

			await Task.Run(async () =>
			{
				try
				{
					using (var sourceCtx = await GetContextAsync(sourceUrl))
					using (var targetCtx = await GetContextAsync(targetUrl))
					{
						List sourceList = sourceCtx.Web.Lists.GetByTitle(sourceTitle);
						List targetList = targetCtx.Web.Lists.GetByTitle(targetListName);

						// Загружаем количество элементов для прогресса
						sourceCtx.Load(sourceList, l => l.ItemCount, l => l.Fields);
						sourceCtx.Load(sourceList.Fields, fs => fs.Include(f => f.InternalName, f => f.ReadOnlyField, f => f.Hidden));
						sourceCtx.ExecuteQuery();

						int totalItems = sourceList.ItemCount;
						int processedCount = 0;

						var fieldsToCopy = sourceList.Fields.AsEnumerable()
							.Where(f => !f.ReadOnlyField && !f.Hidden && 
										f.InternalName != "ContentTypeId" && 
										f.InternalName != "Attachments")
							.ToList();

						CamlQuery query = new CamlQuery();
						query.ViewXml = @"<View Scope='RecursiveAll'>
											<Query><OrderBy><FieldRef Name='ID' Ascending='TRUE' /></OrderBy></Query>
											<RowLimit>25</RowLimit>
										  </View>";

						do
						{
							ct.ThrowIfCancellationRequested(); // Проверка отмены пользователем

							ListItemCollection sourceItems = sourceList.GetItems(query);
							sourceCtx.Load(sourceItems);
							sourceCtx.ExecuteQuery();

							query.ListItemCollectionPosition = sourceItems.ListItemCollectionPosition;

							if (sourceItems.Count > 0)
							{
								foreach (var sourceItem in sourceItems)
								{
									ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
									ListItem newItem = targetList.AddItem(itemCreateInfo);

									foreach (var field in fieldsToCopy)
									{
										try
										{
											object val = sourceItem[field.InternalName];
											if (val != null) newItem[field.InternalName] = val;
										}
										catch { }
									}

									// Копируем системные поля
									newItem["Author"] = sourceItem["Author"];
									newItem["Editor"] = sourceItem["Editor"];
									newItem["Created"] = sourceItem["Created"];
									newItem["Modified"] = sourceItem["Modified"];
									
									// ВАЖНО: Используйте Update(), если права ограничены, 
									// или UpdateOverwriteVersion(), если нужно сохранить даты/авторов (требует прав админа)
									newItem.Update(); 
									processedCount++;
								}

								targetCtx.ExecuteQuery();

								// ОТЧЕТ О ПРОГРЕССЕ
								progress?.Report(new CopyProgressArgs 
								{ 
									Processed = processedCount, 
									Total = totalItems, 
									Message = $"Copied {processedCount} of {totalItems} items..." 
								});
							}
						} while (query.ListItemCollectionPosition != null);
					}
				}
				catch (OperationCanceledException) { throw; } // Пробрасываем отмену выше
				catch (Exception ex)
				{
					System.Diagnostics.Debug.WriteLine($"Ошибка при копировании: {ex.Message}");
					throw;
				}
			});
		}	
		

	}
}