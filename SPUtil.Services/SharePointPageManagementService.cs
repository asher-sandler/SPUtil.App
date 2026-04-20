using Microsoft.SharePoint.Client;
using SPUtil.Infrastructure;
using System;
using System.Threading.Tasks;

namespace SPUtil.Services
{
    /// <summary>
    /// Partial — page management helpers: Exists / Delete / Rename.
    /// Used by the Copy Page workflow in PagesViewModel.
    /// </summary>
    public partial class SharePointService
    {
        // ═══════════════════════════════════════════════════════════════════════
        //  PageExistsAsync
        //  Returns true if a page with the given name exists in the Pages library.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<bool> PageExistsAsync(string siteUrl, string pageName)
        {
            return await Task.Run(async () =>
            {
                string name = pageName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase)
                    ? pageName : pageName + ".aspx";

                using var ctx = await GetContextAsync(siteUrl);
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                await Task.Run(() => ctx.ExecuteQuery());

                string pageRelUrl = ctx.Web.ServerRelativeUrl.TrimEnd('/') + "/Pages/" + name;

                try
                {
                    var file = ctx.Web.GetFileByServerRelativeUrl(pageRelUrl);
                    ctx.Load(file, f => f.Exists);
                    await Task.Run(() => ctx.ExecuteQuery());
                    return file.Exists;
                }
                catch
                {
                    return false;
                }
            });
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  DeletePageAsync
        //  Deletes a Publishing page by name from the Pages library.
        //  Handles CheckOut state — discards checkout before deletion.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task DeletePageAsync(string siteUrl, string pageName)
        {
            await Task.Run(async () =>
            {
                string name = pageName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase)
                    ? pageName : pageName + ".aspx";

                using var ctx = await GetContextAsync(siteUrl);
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                await Task.Run(() => ctx.ExecuteQuery());

                string pageRelUrl = ctx.Web.ServerRelativeUrl.TrimEnd('/') + "/Pages/" + name;

                var file = ctx.Web.GetFileByServerRelativeUrl(pageRelUrl);
                ctx.Load(file, f => f.CheckOutType, f => f.Exists);
                await Task.Run(() => ctx.ExecuteQuery());

                if (!file.Exists)
                {
                    System.Diagnostics.Debug.WriteLine($"[DeletePage] Not found: {pageRelUrl}");
                    return;
                }

                // Discard any pending checkout so deletion is not blocked.
                // If the file has never been checked in (brand-new page),
                // UndoCheckOut throws "no checked-in version" — in that case
                // we skip UndoCheckOut and go straight to DeleteObject.
                if (file.CheckOutType != CheckOutType.None)
                {
                    try
                    {
                        file.UndoCheckOut();
                        await Task.Run(() => ctx.ExecuteQuery());
                    }
                    catch (ServerException ex) when (
                        ex.Message.Contains("no checked") ||
                        ex.Message.Contains("checked in version") ||
                        ex.Message.Contains("Please delete"))
                    {
                        // No checked-in version exists — safe to delete directly
                        System.Diagnostics.Debug.WriteLine(
                            $"[DeletePage] UndoCheckOut skipped (no checked-in version): {ex.Message}");
                    }
                }

                file.DeleteObject();
                await Task.Run(() => ctx.ExecuteQuery());

                System.Diagnostics.Debug.WriteLine($"[DeletePage] Deleted: {pageRelUrl}");
            });
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  RenamePageAsync
        //  Renames a Publishing page by changing its FileLeafRef (filename).
        //  The page Title field is left unchanged.
        //  Used to move an existing page aside before creating a fresh copy.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task RenamePageAsync(string siteUrl, string currentName, string newName)
        {
            await Task.Run(async () =>
            {
                string current = currentName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase)
                    ? currentName : currentName + ".aspx";
                string target = newName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase)
                    ? newName : newName + ".aspx";

                using var ctx = await GetContextAsync(siteUrl);
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                await Task.Run(() => ctx.ExecuteQuery());

                string pageRelUrl = ctx.Web.ServerRelativeUrl.TrimEnd('/') + "/Pages/" + current;

                var file = ctx.Web.GetFileByServerRelativeUrl(pageRelUrl);
                ctx.Load(file, f => f.ListItemAllFields, f => f.CheckOutType);
                await Task.Run(() => ctx.ExecuteQuery());

                // CheckOut is required to change FileLeafRef
                await SafeCheckOutAsync(ctx, file);

                ctx.Load(file.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());

                var item = file.ListItemAllFields;
                item["FileLeafRef"] = target;
                item.Update();
                await Task.Run(() => ctx.ExecuteQuery());

                // CheckIn with minor version — rename should not create a major version
                file.CheckIn($"Renamed from {current} to {target}",
                    CheckinType.MinorCheckIn);
                await Task.Run(() => ctx.ExecuteQuery());

                System.Diagnostics.Debug.WriteLine(
                    $"[RenamePage] {current} → {target}");
            });
        }
    }
}
