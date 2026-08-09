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
        //  Returns true if a page with the given name exists in the Pages library,
        //  optionally inside a subfolder.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<bool> PageExistsAsync(string siteUrl, string pageName, string subfolderPath = "")
        {
            return await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(siteUrl);
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                await Task.Run(() => ctx.ExecuteQuery());

                string pageRelUrl = BuildPageRelativeUrl(
                    ctx.Web.ServerRelativeUrl, pageName, subfolderPath);

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
        //  Deletes a Publishing page by name from the Pages library, optionally
        //  from a subfolder. Handles CheckOut state — discards checkout before
        //  deletion.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task DeletePageAsync(string siteUrl, string pageName, string subfolderPath = "")
        {
            await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(siteUrl);
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                await Task.Run(() => ctx.ExecuteQuery());

                string pageRelUrl = BuildPageRelativeUrl(
                    ctx.Web.ServerRelativeUrl, pageName, subfolderPath);

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
        //  The page Title field is left unchanged, and the page stays in its
        //  current folder.
        //  Used to move an existing page aside before creating a fresh copy.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task RenamePageAsync(string siteUrl, string currentName, string newName, string subfolderPath = "")
        {
            await Task.Run(async () =>
            {
                // Only the file name goes into FileLeafRef — the folder never changes,
                // so the subfolder is not part of the target name.
                string target = newName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase)
                    ? newName : newName + ".aspx";

                using var ctx = await GetContextAsync(siteUrl);
                ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
                await Task.Run(() => ctx.ExecuteQuery());

                string pageRelUrl = BuildPageRelativeUrl(
                    ctx.Web.ServerRelativeUrl, currentName, subfolderPath);

                var file = ctx.Web.GetFileByServerRelativeUrl(pageRelUrl);
                // Level is read BEFORE checking out — afterwards it always reports
                // Checkout and the original publication state would be lost.
                ctx.Load(file, f => f.ListItemAllFields, f => f.CheckOutType, f => f.Level);
                await Task.Run(() => ctx.ExecuteQuery());

                FileLevel originalLevel = file.Level;
                System.Diagnostics.Debug.WriteLine(
                    $"[RenamePage] Original level of {pageRelUrl}: {originalLevel}");

                // CheckOut is required to change FileLeafRef
                await SafeCheckOutAsync(ctx, file);

                ctx.Load(file.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());

                var item = file.ListItemAllFields;
                item["FileLeafRef"] = target;
                item.Update();
                await Task.Run(() => ctx.ExecuteQuery());

                // The rename invalidated the object path of `file`: it still points at
                // the old URL, which no longer exists. CSOM re-resolves object paths in
                // every batch, so calling CheckIn on it fails with 'Unknown Error'.
                // Re-acquire the file under its new name before checking in.
                string renamedRelUrl = BuildPageRelativeUrl(
                    ctx.Web.ServerRelativeUrl, target, subfolderPath);
                var renamedFile = ctx.Web.GetFileByServerRelativeUrl(renamedRelUrl);

                // Restore the publication state the page had before the rename.
                // Renaming is a technical operation and must not silently take a
                // published page offline, nor publish a draft the user never approved.
                if (originalLevel == FileLevel.Published)
                {
                    renamedFile.CheckIn($"Renamed from {currentName} to {target}",
                        CheckinType.MajorCheckIn);
                    await Task.Run(() => ctx.ExecuteQuery());

                    try
                    {
                        renamedFile.Publish($"Republished after rename to {target}");
                        await Task.Run(() => ctx.ExecuteQuery());
                    }
                    catch (ServerException ex)
                    {
                        // A major check-in already publishes the file in libraries
                        // without content approval — Publish then reports that there
                        // is nothing to publish. Harmless.
                        System.Diagnostics.Debug.WriteLine(
                            $"[RenamePage] Publish skipped: {ex.Message}");
                    }
                }
                else
                {
                    // Draft, or the page was already checked out when we started and
                    // the pre-checkout state is unknowable — stay minor, which is the
                    // safer of the two.
                    renamedFile.CheckIn($"Renamed from {currentName} to {target}",
                        CheckinType.MinorCheckIn);
                    await Task.Run(() => ctx.ExecuteQuery());
                }

                System.Diagnostics.Debug.WriteLine(
                    $"[RenamePage] {pageRelUrl} → {renamedRelUrl} (level {originalLevel})");
            });
        }
        // ═══════════════════════════════════════════════════════════════════════
        //  Private helpers
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Builds the server-relative URL of a page inside the Pages library,
        /// honouring an optional subfolder. The subfolder may be nested to any
        /// depth ("Admin" or "FacultyAdmin/Sub/Deeper"); an empty value means the
        /// library root.
        /// </summary>
        private static string BuildPageRelativeUrl(
            string webServerRelativeUrl,
            string pageName,
            string subfolderPath)
        {
            string name = pageName.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase)
                ? pageName : pageName + ".aspx";

            string root = webServerRelativeUrl.TrimEnd('/') + "/Pages";

            // Accept both slash styles and tolerate leading/trailing separators.
            // Inner separators are preserved, so nested paths work unchanged.
            string sub = (subfolderPath ?? string.Empty).Replace('\\', '/').Trim('/');

            return string.IsNullOrEmpty(sub)
                ? $"{root}/{name}"
                : $"{root}/{sub}/{name}";
        }
    }
}