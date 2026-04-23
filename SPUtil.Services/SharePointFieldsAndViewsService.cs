using Microsoft.SharePoint.ApplicationPages.MetaWeblog;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
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
using Serilog;
//using SPUtil.Infrastructure;

namespace SPUtil.Services
{
	public partial class SharePointService : ISharePointService
	{
		
       public async Task<List<FieldInfo>> GetFieldInfosFromSiteAsync(string siteUrl, string listTitle)
        {
            var fieldInfos = new List<FieldInfo>();

            using (var ctx = await GetContextAsync(siteUrl))
            {
                // Настройте Credentials, если это необходимо
                Web web = ctx.Web;
                List list = web.Lists.GetByTitle(listTitle);

                // 1. Первичная загрузка коллекции полей с базовыми свойствами
                ctx.Load(list.Fields, fs => fs.Include(
                    f => f.Title,
                    f => f.InternalName,
                    f => f.StaticName,
                    f => f.FieldTypeKind,
                    f => f.Required,
                    f => f.ReadOnlyField,
                    f => f.Hidden,
                    f => f.SchemaXml,
                    f => f.DefaultValue // Загружаем значение по умолчанию сразу
                ));

                await Task.Run(() => ctx.ExecuteQuery());

                foreach (var field in list.Fields)
                {

                    System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] Field Type: '{field.FieldTypeKind}' .");

                    string[] emailFields = { "EmailSender", "EmailTo", "EmailCc", "EmailFrom", "EmailSubject", "EmailHeaders" };

					// Если поле скрытое, но оно входит в список Email-полей — НЕ пропускаем его
					if (emailFields.Contains(field.InternalName))
						continue;
					
                    // --- ЛОГИКА ФИЛЬТРАЦИИ ---
                    if (field.Hidden || (field.InternalName.StartsWith("_") && !field.InternalName.Contains("_x")) || field.InternalName.StartsWith("vti_"))
                        continue;

                    // Пропускаем ReadOnly, кроме вычисляемых полей
                    if (    field.ReadOnlyField &&
                            field.FieldTypeKind != FieldType.Calculated &&
                            field.FieldTypeKind != FieldType.Lookup)
                    {
                        continue;
                    }
                    // Список системных исключений
                    string[] systemExclusions =
                    {
                "ContentType", "Attachments", "FolderChildCount", "ItemChildCount","ParentLeafName","ParentVersionString",
                "Edit", "LinkTitle", "Order", "GUID", "AppAuthor", "AppEditor", "DocIcon", "FileLeafRef","Title"
            };
                    if (systemExclusions.Contains(field.InternalName)) continue;

                    // Создаем объект и заполняем базовые данные
                    var info = new FieldInfo
                    {
                        Name = field.InternalName,
                        DisplayName = field.Title,
                        StaticName = field.StaticName,
                        FieldType = field.FieldTypeKind.ToString(),
                        Required = field.Required ? "TRUE" : "FALSE",
                        DefaultValue = field.DefaultValue ?? string.Empty,
                        SchemaXml = field.SchemaXml
                    };

                    // --- ОБРАБОТКА СПЕЦИФИЧЕСКИХ СВОЙСТВ (Как в PowerShell) ---
                    // Для каждого типа делаем Cast и Load специфических свойств, затем ExecuteQuery
                    try
                    {
                        switch (field.FieldTypeKind)
                        {
                            case FieldType.Text:
                                var textField = ctx.CastTo<FieldText>(field);
                                ctx.Load(textField, f => f.MaxLength);
                                ctx.ExecuteQuery();
                                info.MaxLength = textField.MaxLength;
                                break;

                            case FieldType.URL:
                                var urlField = ctx.CastTo<FieldUrl>(field);
                                ctx.Load(urlField, f => f.DisplayFormat);
                                ctx.ExecuteQuery();
                                // ТЕПЕРЬ ТУТ БУДЕТ "Hyperlink" или "Image"
                                info.Format = urlField.DisplayFormat.ToString();
                                break;

                            case FieldType.DateTime:
                                var dateField = ctx.CastTo<FieldDateTime>(field);
                                ctx.Load(dateField, f => f.DisplayFormat);
                                ctx.ExecuteQuery();
                                // Будет "DateOnly" или "DateTime"
                                info.Format = dateField.DisplayFormat.ToString();
                                break;

                            case FieldType.Choice:
                            case FieldType.MultiChoice:
                                var choiceField = ctx.CastTo<FieldChoice>(field);
                                ctx.Load(choiceField, f => f.Choices, f => f.EditFormat);
                                ctx.ExecuteQuery();
                                info.Choices = choiceField.Choices.ToList();
                                info.Format = choiceField.EditFormat.ToString(); // Dropdown/RadioButtons
                                break;

							   case FieldType.Lookup:
								var lookupField = ctx.CastTo<FieldLookup>(field);
								
								// 1. ОБЯЗАТЕЛЬНО добавляем f => f.PrimaryFieldId в Load
								ctx.Load(lookupField, 
									f => f.LookupList, 
									f => f.LookupField, 
									f => f.PrimaryFieldId); // Без этого будет ошибка инициализации
								
								ctx.ExecuteQuery();

								info.LookupListId = lookupField.LookupList;
								info.LookupFieldName = lookupField.LookupField;

                                // 2. Проверяем на зависимый Lookup (например, sort1:Code)
                                
                                if (Guid.TryParse(lookupField.PrimaryFieldId, out Guid x))
                                {
                                    // Только если это Guid, мы конвертируем его в строку для вашей модели
                                    info.PrimaryFieldId = lookupField.PrimaryFieldId.ToString();
                                    info.IsDependentLookup = true;
                                    System.Diagnostics.Debug.WriteLine($"[INFO] Dependent field {field.InternalName} found. Parent ID: {info.PrimaryFieldId}");
                                }
                                else
                                {
                                    info.PrimaryFieldId = string.Empty;
                                    info.IsDependentLookup = false;
                                }


								// Логика получения имени целевого списка (для основных лукапов)
								if (Guid.TryParse(lookupField.LookupList, out Guid g))
								{
						            try
						            {
							            var targetList = ctx.Web.Lists.GetById(g);
							            ctx.Load(targetList, l => l.Title);
							            ctx.ExecuteQuery();
							            info.LookupListName = targetList.Title;
						            }
						            catch { /* List may be on another web or deleted */ }
						                _log.Error("ERROR in catch block");
					            }
					
					break;                            
					case FieldType.Calculated:
                                var calcField = ctx.CastTo<FieldCalculated>(field);
                                // Загружаем OutputType и DateFormat
                                ctx.Load(calcField, f => f.Formula, f => f.OutputType, f => f.DateFormat);
                                ctx.ExecuteQuery();

                                info.Formula = calcField.Formula;
                                info.ResultType = calcField.OutputType.ToString();

                                // Если результат — дата, сохраняем формат (DateOnly или DateTime)
                                if (calcField.OutputType == FieldType.DateTime)
                                {
                                    info.Format = calcField.DateFormat.ToString();
                                }
                                // Если формат всё равно пустой, пробуем вытащить напрямую из XML
                                if (string.IsNullOrEmpty(info.Format))
                                {
                                    info.Format = _cloneService.GetAttributeFromXml(field.SchemaXml, "Format");
                                }
                                break;
                            case FieldType.Note:
                                var noteField = ctx.CastTo<FieldMultiLineText>(field);
                                // Загружаем только то, что точно есть (NumberOfLines и RichText обычно присутствуют везде)
                                ctx.Load(noteField, f => f.NumberOfLines, f => f.RichText);
                                ctx.ExecuteQuery();

                                info.NumLines = noteField.NumberOfLines;
                                info.RichText = noteField.RichText ? "TRUE" : "FALSE";

                                // Вместо AllowFullRichText используем парсинг SchemaXml, 
                                // чтобы достать RichTextMode (Compatible, Html, FullHtml)
                                info.RichTextMode = _cloneService.GetAttributeFromXml(field.SchemaXml, "RichTextMode") ?? "Compatible";

                                // Также можно достать IsolateStyles
                                info.IsolateStyles = _cloneService.GetAttributeFromXml(field.SchemaXml, "IsolateStyles") ?? "FALSE";
                                break;
                            case FieldType.Number:
                            case FieldType.Currency:
                                var numField = ctx.CastTo<FieldNumber>(field);
                                // Убираем DisplayFormat, оставляем Min/Max
                                ctx.Load(numField, f => f.MinimumValue, f => f.MaximumValue);
                                ctx.ExecuteQuery();

                                info.MinValue = numField.MinimumValue;
                                info.MaxValue = numField.MaximumValue;

                                // Если вам нужны знаки после запятой (Decimals), 
                                // их лучше вытащить из SchemaXml парсингом, так как свойства в SDK часто нет
                                break;
                            case FieldType.Invalid:
                                if (field.TypeAsString == "TargetTo")
                                {
                                    // Для этого типа нам не нужно загружать доп. свойства через CastTo,
                                    // так как все настройки сидят в данных айтемов, а не в схеме поля.
                                    info.FieldType = "TargetTo";
                                    System.Diagnostics.Debug.WriteLine($"[SP_SERVICE] Specialized field detected: {field.InternalName} (TargetTo)");
                                }
                                break;
                        }
                    }
                    catch (Exception ex)
                    {
                        _log.Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                        // Если по какому-то полю не удалось получить детали, 
                        // логируем, но не прерываем весь процесс
                        System.Diagnostics.Debug.WriteLine($"Error fetching details for field {field.InternalName}: {ex.Message}");
                    }

                    fieldInfos.Add(info);
                }
            }
			//foreach(var fXml in fieldInfos){
				//System.Diagnostics.Debug.WriteLine($"Fields schema: {fXml.BuildXml()}");
			//}
            return fieldInfos;
        }

		
		public async Task<List<SPFieldData>> GetListFieldsAsync(string siteUrl, string listPath)
		{
			// Запускаем синхронный код CSOM в отдельном потоке
			return await Task.Run(async () =>
			{
				//_spService.NormalizeUrl(RightSiteUrl)
				using (var context = await GetContextAsync(siteUrl))
                {
					//context.Credentials = GetCredentials();

					// Парсим GUID (ваш 100149a3-951e-4e22-aedd-0ad603a3c99a)
					if (!Guid.TryParse(listPath, out Guid listId))
						throw new Exception("Invalid list GUID format");

					List list = context.Web.Lists.GetById(listId);

					// Загружаем поля
					var fields = list.Fields;
					context.Load(fields, col => col.Include(
								f => f.Id,
								f => f.Title,
								f => f.InternalName,
								f => f.TypeAsString,
								f => f.Hidden,
								f => f.FieldTypeKind,
								f => f.FromBaseType)); // Добавили FromBaseType и Hidden
													   // В версии 15.0 используем СИНХРОННЫЙ метод
					context.ExecuteQuery();
					
                    System.Diagnostics.Debug.WriteLine($"--- Field list for list {listId} ---");

                    foreach (var field in fields)
                    {
                        // Выводим название и внутреннее имя каждого поля
                        System.Diagnostics.Debug.WriteLine($"Field: {field.Title} | InternalName: {field.InternalName} | Type: {field.FieldTypeKind}");
                    }

                    System.Diagnostics.Debug.WriteLine($"--- Total fields: {fields.Count} ---");
					
					// 1. Сначала получаем ВСЕ поля, которые прошли базовый технический фильтр
					var allFieldsList = fields.ToList();

					// 2. Список системных имен, которые мы хотим видеть ВСЕГДА (даже если они похожи на системные)
					var priorityIds = new List<string> { "ID", "Title" };
					var endSystemNames = new List<string> { "Created", "Author", "Modified", "Editor" };

					var filteredResult = allFieldsList.Where(f =>
					{
						string name = f.InternalName;

						// УСЛОВИЕ ДЛЯ ИВРИТА: 
						// Если имя начинается на _x (Unicode), это пользовательское поле на иврите/арабском и т.д.
						// Мы ДОЛЖНЫ его оставить.
						if (name.StartsWith("_x")) return true;

						// Пропускаем скрытые поля, кроме ID (он часто Hidden в некоторых списках)
						if (f.Hidden && !name.Equals("ID", StringComparison.OrdinalIgnoreCase)) return false;

						// Отсекаем технические вычисляемые поля (но не Calculated поля пользователя)
						if (f.TypeAsString == "Computed") return false;

						// Отсекаем настоящие системные поля на "_" (которые НЕ иврит)
						if (name.StartsWith("_") && !priorityIds.Contains(name)) return false;

						// Список исключений (то, что не хотим видеть в таблице)
						var blacklist = new HashSet<string>(StringComparer.OrdinalIgnoreCase) 
						{ 
							"ContentTypeId", "Attachments", "Edit", "DocIcon", 
							"AppAuthor", "ItemChildCount", "FolderChildCount", "AppEditor", "vti_folderitemcount"
						};

						if (blacklist.Contains(name)) return false;

						return true;
					}).ToList();

					// 3. СОРТИРОВКА: ID, Title -> Остальные по Алфавиту -> Системные в конце
					var startFields = filteredResult
						.Where(f => priorityIds.Contains(f.InternalName))
						.OrderBy(f => priorityIds.IndexOf(f.InternalName));

					var endFields = filteredResult
						.Where(f => endSystemNames.Contains(f.InternalName))
						.OrderBy(f => endSystemNames.IndexOf(f.InternalName));

					var middleFields = filteredResult
						.Where(f => !priorityIds.Contains(f.InternalName) && !endSystemNames.Contains(f.InternalName))
						.OrderBy(f => f.Title); // Здесь поля на иврите будут отсортированы по их отображаемому Title

					return startFields
						.Concat(middleFields)
						.Concat(endFields)
						.Select(f => new SPFieldData
						{
							Id = f.Id.ToString(),
							Title = f.Title, // Здесь будет нормальный текст на иврите
							InternalName = f.InternalName, // Здесь будет _x05...
							TypeAsString = f.TypeAsString
						})
						.ToList();					

				}
			});
		}
		public async Task<List<SPViewData>> GetListViewsAsync(string siteUrl, string listPath)
		{
			return await Task.Run(async () =>
			{
				var viewDataList = new List<SPViewData>();
				using (var ctx =   await GetContextAsync(siteUrl))
				{
					// Configure Credentials here (Credentials = ...)
					
					Microsoft.SharePoint.Client.Web web = ctx.Web;

                    Microsoft.SharePoint.Client.List list;

                    if (Guid.TryParse(listPath, out Guid guid))
                    {
                        list = web.Lists.GetById(new Guid(listPath));
                    }
                    else
                    {
                        list = web.Lists.GetByTitle(listPath);
                    }
                    //(listPath.Length > 30) 
					//	? web.Lists.GetById(new Guid(listPath)) 
					//	: 

					// Добавляем ListViewXml в загрузку
					ctx.Load(list.Views, vs => vs.Include(
						v => v.Title,
						v => v.Id,
						v => v.ViewQuery,
						v => v.DefaultView,
						v => v.ViewFields,
						v => v.ServerRelativeUrl,
                        v => v.Hidden,
						v => v.ListViewXml,
						v => v.Aggregations // Загружаем агрегации
					));
					ctx.ExecuteQuery();

					foreach (var v in list.Views)
					{
                        if (!v.Hidden)
                        {
                            viewDataList.Add(new SPViewData
                            {
                                Title = v.Title,
                                Id = v.Id.ToString(),
                                ViewQuery = v.ViewQuery,
                                DefaultView = v.DefaultView,
                                ViewFields = v.ViewFields.ToArray(),
                                ServerRelativeUrl = v.ServerRelativeUrl,
                                Aggregations = v.Aggregations, // Сохраняем агрегации
                                SchemaXml = v.ListViewXml // Сохраняем схему для клонирования
                            });
                        }
					}
				}
				return viewDataList;
			});
		}		
        public async Task<List<string>> GetListViewSchemasAsync(string siteUrl, string listTitle)
		{
			return await Task.Run(async () =>
			{
				var viewSchemas = new List<string>();
				using (var ctx =  await GetContextAsync(siteUrl))

                {
					// Настройка Credentials
					var list = ctx.Web.Lists.GetByTitle(listTitle);
					var views = list.Views;
					ctx.Load(views, vs => vs.Include(v => v.ListViewXml, v => v.Hidden));
					ctx.ExecuteQuery();

					foreach (var view in views)
					{
						if (!view.Hidden) viewSchemas.Add(view.ListViewXml);
					}
				}
				return viewSchemas;
			});
		}			

	}
}