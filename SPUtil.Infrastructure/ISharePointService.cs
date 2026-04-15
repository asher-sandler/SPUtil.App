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
    }
}