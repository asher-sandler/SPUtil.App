using System.Collections.Generic; // Для List<>
using System.Collections.ObjectModel;
using System.Threading.Tasks;
using SPUtil.Infrastructure;
//using Microsoft.SharePoint.Client;

//using SPUtil.UsingUtils;

namespace SPUtil.Services
{
    public interface ISharePointService
    {
		string GetCurrentUsername();
		//string NormalizeUrl(string url);
		//string UrlWithF5(string url);
		//string GetConnectionStatus();
		
		
        Task<ObservableCollection<SPNode>> GetSiteStructureAsync(string siteUrl);
        Task<string> GetDetailedInfoAsync(string siteUrl, string listId, int templateId);
        //Task CopyListAsync(string sourceUrl, string targetUrl, string listName);

        Task<List<SPListItemData>> GetListItemsByIDAsync(string siteUrl, string listId);
		Task<string> GetListNameByIdAsync(string siteUrl, string listId);
		Task<Guid> GetListIdByTitleAsync(string siteUrl, string listTitle);

		Task DeleteListAsync(string siteUrl, string listTitle);
        Task<List<SPListItemData>> GetListItemsByTitleAsync(
			string siteUrl, 
			string listTitle, 
			IProgress<int> progress, 
			CancellationToken ct);
			
		Task CopyFolderStructureAsync(
					string sourceUrl,
					string targetUrl,
					string sourceLibraryTitle,
					string targetLibraryTitle,
					IProgress<string> progress = null);

        Task CopyDocLibFilesAsync(
			string sourceUrl,
			string targetUrl,
			string sourceLibraryTitle,
			string targetLibraryTitle,
			string action,                       // "Append" | "Overwrite" | "Mirror"
			IProgress<CopyProgressArgs> progress,
			CancellationToken ct);

        Task<bool> ListExistsAsync(string siteUrl, string listTitle);
		Task<List<FieldInfo>> GetFieldInfosFromSiteAsync(string siteUrl, string listTitle);
		Task<AuthResult> ValidateConnectionAsync(string siteUrl);



        Task<List<SPFieldData>> GetListFieldsAsync(string siteUrl, string listPath);
        Task<List<SPViewData>> GetListViewsAsync(string siteUrl, string listPath);
        Task<List<SPFileData>> GetLibraryItemsAsync(string siteUrl, string listId);
        Task<List<SPFileData>> GetPageItemsAsync(string siteUrl, string listId);
        Task<List<SPWebPartData>> GetWebPartsAsync(string siteUrl, string fileRelativeUrl);
		
		Task<List<string>> GetListSchemaAsync(string siteUrl, string listTitle);

		Task<string> GetListTitleByGuidAsync(string siteUrl, Guid listGuid);
		//Task CreateListFromSchemaAsync(string targetUrl, string listTitle, List<string> fieldSchemas);				
		Task<List<string>> GetListViewSchemasAsync(string siteUrl, string listTitle);
		Task CreateDocLibAsync(string siteUrl, string listName, string displayName = "");

		Task CreateListFromSchemaAsync(string targetUrl, string internalListName, string newlistTitle, List<FieldInfo> sourceFields, List<SPViewData> sourceViews, int listType);


        Task CopyListItemsAsync(
			string sourceUrl, 
			string targetUrl, 
			string sourceTitle, 
			string targetListName, // Новый параметр
			string action,         // "Overwrite" или "Append"
			IProgress<CopyProgressArgs> progress, 
			CancellationToken ct);
			
		Task ClearListItemsAsync(string siteUrl, string listTitle);
        Task<SPListInfo> GetListDetailedInfoAsync(string siteUrl, string listTitle);

		// ═══════════════════════════════════════════════════════════════════════════════
		//  ISharePointService — расширение для операций со страницами и WebParts
		//  Добавить эти методы в существующий интерфейс ISharePointService.cs
		// ═══════════════════════════════════════════════════════════════════════════════

		// ── 1. Чтение ────────────────────────────────────────────────────────────────

		/// <summary>
		/// Читает полный снимок страницы: Layout, PublishingHtml, все WebParts
		/// с их StorageKey, ZoneKey, ExportXml и Properties в визуальном порядке.
		/// </summary>
		Task<PageSnapshot> GetPageSnapshotAsync(string siteUrl, string pageRelativeUrl);

		// ── 2. Создание ──────────────────────────────────────────────────────────────

		/// <summary>
		/// Создаёт новую Publishing-страницу и воспроизводит на ней все WebParts
		/// из снимка с сохранением визуального порядка и всех свойств.
		/// </summary>
		Task CreatePageFromSnapshotAsync(
			string targetSiteUrl,
			string targetPageName,       // без .aspx
			PageSnapshot snapshot);

		// ── 3. Изменить один WebPart ─────────────────────────────────────────────────

		/// <summary>
		/// Обновляет свойства одного WebPart на странице.
		/// Поддерживает изменение Title, произвольных свойств и позиции.
		/// </summary>
		Task UpdateWebPartAsync(
			string siteUrl,
			string pageRelativeUrl,
			WebPartUpdateRequest request);

		// ── 4. Изменить все WebParts ─────────────────────────────────────────────────

		/// <summary>
		/// Применяет список изменений ко всем указанным WebParts за одну операцию
		/// CheckOut / CheckIn.
		/// </summary>
		Task UpdateAllWebPartsAsync(
			string siteUrl,
			string pageRelativeUrl,
			IEnumerable<WebPartUpdateRequest> requests);

		// ── 5. Добавить WebPart ──────────────────────────────────────────────────────

		/// <summary>
		/// Добавляет WebPart на страницу из его ExportXml.
		/// Регистрирует в wpz и вставляет заглушку в PublishingContent
		/// на указанную позицию (position=0 — в конец).
		/// Возвращает StorageKey добавленного WebPart.
		/// </summary>
		Task<string> AddWebPartAsync(
			string siteUrl,
			string pageRelativeUrl,
			string webPartXml,
			int position = 0);            // 0 = append

		// ── 6. Удалить WebPart ───────────────────────────────────────────────────────

		/// <summary>
		/// Удаляет WebPart со страницы по StorageKey:
		/// удаляет объект из wpz и убирает заглушку из PublishingContent.
		/// </summary>
		Task DeleteWebPartAsync(
			string siteUrl,
			string pageRelativeUrl,
			string storageKey);

		// ── 7. Изменить одно свойство ────────────────────────────────────────────────

		/// <summary>
		/// Точечно обновляет одно свойство WebPart без затрагивания остальных.
		/// Shortcut для UpdateWebPartAsync с одним свойством.
		/// </summary>
		Task UpdateWebPartPropertyAsync(
			string siteUrl,
			string pageRelativeUrl,
			string storageKey,
			string propertyName,
			string propertyValue);

		// ── 8. Изменить порядок WebParts ─────────────────────────────────────────────

		/// <summary>
		/// Переставляет WebParts в PublishingContent согласно новому порядку StorageKey.
		/// Сами объекты WebPart не затрагиваются — только перестраивается HTML.
		/// </summary>
		Task ReorderWebPartsAsync(
			string siteUrl,
			string pageRelativeUrl,
			IEnumerable<string> orderedStorageKeys);  // в нужном порядке

		// ── 9. Переместить WebPart между страницами ───────────────────────────────────

		/// <summary>
		/// Перемещает WebPart с исходной страницы на целевую:
		/// читает ExportXml → добавляет на target → удаляет с source.
		/// </summary>
		Task MoveWebPartAsync(
			string siteUrl,
			string sourcePageRelativeUrl,
			string targetPageRelativeUrl,
			string storageKey,
			int targetPosition = 0);

		// ── 10. Клонировать WebPart внутри страницы ───────────────────────────────────

		/// <summary>
		/// Добавляет копию WebPart на ту же страницу с теми же свойствами
		/// но новым StorageKey. Возвращает StorageKey новой копии.
		/// </summary>
		Task<string> CloneWebPartAsync(
			string siteUrl,
			string pageRelativeUrl,
			string storageKey,
			int targetPosition = 0);

		// ── 11. Получить ExportXml одного WebPart ────────────────────────────────────

		/// <summary>
		/// Скачивает .webpart XML через exportwp.aspx для указанного WebPart.
		/// Используется внутри других методов и может быть вызван напрямую.
		/// </summary>
		Task<string> GetWebPartExportXmlAsync(
			string siteUrl,
			string pageRelativeUrl,
			string storageKey);

		// ── 12. Восстановить страницу из снимка ──────────────────────────────────────

		/// <summary>
		/// Применяет снимок к уже существующей странице:
		/// удаляет все текущие WebParts и воспроизводит WebParts из снимка.
		/// В отличие от CreatePageFromSnapshotAsync — не создаёт страницу.
		/// </summary>
		Task RestorePageFromSnapshotAsync(
			string siteUrl,
			string pageRelativeUrl,
			PageSnapshot snapshot);

		// ── 13. Сравнить два снимка ───────────────────────────────────────────────────

		/// <summary>
		/// Сравнивает два снимка страниц.
		/// Возвращает текстовый diff: какие WebParts добавились,
		/// удалились или изменили свойства.
		/// </summary>
		Task<string> ComparePageSnapshotsAsync(
			PageSnapshot source,
			PageSnapshot target);

		// ── 14. CheckOut / CheckIn / Publish как отдельные методы ────────────────────

		Task CheckOutPageAsync(string siteUrl, string pageRelativeUrl);

		Task CheckInPageAsync(string siteUrl, string pageRelativeUrl, string comment = "");

		Task PublishPageAsync(string siteUrl, string pageRelativeUrl, string comment = "");

		// ── 15. Получить все страницы с WebParts ──────────────────────────────────────

		/// <summary>
		/// Читает снимки всех страниц из библиотеки Pages указанного сайта.
		/// Используется для массового клонирования или аудита.
		/// </summary>
		Task<List<PageSnapshot>> GetAllPagesSnapshotsAsync(string siteUrl);


    }
}