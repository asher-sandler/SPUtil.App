using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Publishing;
using Microsoft.SharePoint.Client.WebParts;
using SPUtil.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using Serilog;
using HtmlAgilityPack;
using System.Threading.Tasks;
using System.Net.Http;

namespace SPUtil.Services
{
    /// <summary>
    /// Partial class — операции со страницами и WebParts.
    /// Все методы используют GetContextAsync и GetCredentials из основного файла.
    /// </summary>
    public partial class SharePointService
    {
        //private static readonly ILogger _log = Log.ForContext("SourceContext", "SharePointPageService");

        private static readonly ILogger _logPage = Log.ForContext("SourceContext", "SharePointPageService");

        // ═══════════════════════════════════════════════════════════════════════
        //  ВНУТРЕННИЕ ВСПОМОГАТЕЛЬНЫЕ МЕТОДЫ
        // ═══════════════════════════════════════════════════════════════════════

        /// <summary>
        /// Скачивает рендеренный HTML страницы через HTTP GET.
        /// Используется для парсинга webpartid / webpartid2.
        /// </summary>
        private async Task<string> FetchPageHtmlAsync(string siteUrl, string pageRelativeUrl)
        {
            string hostRoot = "https://" + new Uri(siteUrl).Host;
            string fullUrl  = hostRoot + pageRelativeUrl;

            // NormalizeUrl removes the trailing "2" from the host (portals2 → portals)
            // so the request goes through the correct network path instead of the proxy.
            fullUrl = SPUtil.Infrastructure.SPUsingUtils.NormalizeUrl(fullUrl);

            var handler = new System.Net.Http.HttpClientHandler
            {
                Credentials = GetCredentials(siteUrl)
            };
            using var http = new System.Net.Http.HttpClient(handler);
            http.DefaultRequestHeaders.Add("Accept", "text/html");

            var response = await http.GetAsync(fullUrl);
            response.EnsureSuccessStatusCode();
            return await response.Content.ReadAsStringAsync();
        }

        /// <summary>
        /// Парсит рендеренный HTML страницы и возвращает словарь
        /// ZoneKey (webpartid2) → StorageKey (webpartid).
        /// Атрибуты берутся из div вида:
        ///   webpartid="StorageKey" webpartid2="ZoneKey"
        /// </summary>
        private Dictionary<string, string> ParseZoneKeyToStorageKey(string renderedHtml)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Match: webpartid="GUID" ... webpartid2="GUID"
            // Both attributes appear on the same div element
            var matches = Regex.Matches(
                renderedHtml,
                @"webpartid=""([0-9a-f\-]{36})""[^>]*webpartid2=""([0-9a-f\-]{36})""",
                RegexOptions.IgnoreCase);

            foreach (Match m in matches)
            {
                string storageKey = m.Groups[1].Value.ToLower();
                string zoneKey    = m.Groups[2].Value.ToLower();
                result[zoneKey]   = storageKey;
            }

            return result;
        }

		private string ReplaceFailedWebPartPlaceholders(
			string html, IEnumerable<WebPartCopyEntry> failedEntries)
		{
			var doc = new HtmlDocument();
			doc.LoadHtml(html);

			foreach (var entry in failedEntries)
			{
				if (string.IsNullOrEmpty(entry.SourceZoneKey)) continue;

				var innerDiv = doc.DocumentNode.SelectSingleNode(
					$"//div[contains(@class,'ms-rtestate-read') and contains(@class,'{entry.SourceZoneKey}')]");

				var outerDiv = innerDiv?.ParentNode;
				if (outerDiv == null)
				{
					_logPage.Caller().Warning(
						"ReplaceFailedWebPartPlaceholders: placeholder not found for '{Title}' (ZoneKey={ZoneKey})",
						entry.Title, entry.SourceZoneKey);
					continue;
				}

				try
				{
					var warningNode = HtmlNode.CreateNode(
						$"<div>" +
						$"<p>Web Part <b> \"{System.Net.WebUtility.HtmlEncode(entry.Title)}\"</b> not available on this site and cannot be added to this page.<br/>" +
						$"Please install/activate the corresponding feature on this site and copy the page again, " +
						$"or contact IT support: <a href=\"mailto:supportsp@savion.huji.ac.il\">supportsp@savion.huji.ac.il</a></p>" +
						$"<p style=\"font-size:11px;color:#666;\">Technical details: {System.Net.WebUtility.HtmlEncode(entry.Reason)}</p>" +
						$"</div>");

					outerDiv.ParentNode.ReplaceChild(warningNode, outerDiv);
				}
				catch (Exception ex)
				{
					// A single placeholder failing to render must not abort the whole
					// page creation — same principle already applied to the zone case.
					_logPage.Caller().Error(ex,
						"ReplaceFailedWebPartPlaceholders: could not build/insert warning node for '{Title}' (ZoneKey={ZoneKey})",
						entry.Title, entry.SourceZoneKey);
				}
			}

			return doc.DocumentNode.OuterHtml;
		}
        /// <summary>
        /// Извлекает ZoneKey GUID-ы из PublishingPageContent HTML
        /// в визуальном порядке (порядок появления в тексте).
        /// Ищет: ms-rtestate-read {GUID} в class атрибуте div.
        /// </summary>
        private List<string> ParseZoneKeysInOrder(string publishingHtml)
        {
            var result  = new List<string>();
            var matches = Regex.Matches(
                publishingHtml,
                @"ms-rtestate-read\s+([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})",
                RegexOptions.IgnoreCase);

            foreach (Match m in matches)
            {
                string guid = m.Groups[1].Value.ToLower();
                if (!result.Contains(guid))
                    result.Add(guid);
            }
            return result;
        }

        /// <summary>
        /// Builds an ms-rte-wpbox placeholder div for PublishingPageContent.
        /// If sourceMeta is provided — inserts an SPUTIL JSON comment inside the div
        /// so SyncPropertiesAsync can later find and process it.
        /// </summary>
        private static string BuildWpBoxPlaceholder(
            string zoneKey,
            WebPartPlaceholderMeta? sourceMeta = null)
        {
            string metaComment = string.Empty;
            string extraClass  = string.Empty;

            if (sourceMeta != null)
            {
                // Escape values for safe JSON embedding
                string safeTitle   = sourceMeta.Title.Replace("\"", "\\\"").Replace("\n", " ");
                string safeSiteUrl = sourceMeta.SiteUrl.Replace("\"", "\\\"");
                string safePageUrl = sourceMeta.PageUrl.Replace("\"", "\\\"");
                string safeZoneId  = sourceMeta.ZoneId.Replace("\"", "\\\"");

                string json =
                    $"{{" +
                    $"\"storageKey\":\"{sourceMeta.StorageKey}\"," +
                    $"\"title\":\"{safeTitle}\"," +
                    $"\"position\":{sourceMeta.Position}," +
                    $"\"zoneId\":\"{safeZoneId}\"," +
                    $"\"siteUrl\":\"{safeSiteUrl}\"," +
                    $"\"pageUrl\":\"{safePageUrl}\"" +
                    $"}}";

                metaComment = $"\r\n  <!--SPUTIL:{json}-->" +
                              $"\r\n  <p style=\"border:2px dashed #cc4400;padding:8px;" +
                              $"background:#fff8f0;font-family:Consolas;font-size:12px;margin:4px 0\">" +
                              $"[WebPart Placeholder] <b>{System.Web.HttpUtility.HtmlEncode(sourceMeta.Title)}</b>" +
                              $" &nbsp;|&nbsp; Position: {sourceMeta.Position}" +
                              $" &nbsp;|&nbsp; Zone: {sourceMeta.ZoneId}<br/>" +
                              $"<i>Add this WebPart manually, then run Sync Properties to restore its settings.</i>" +
                              $"</p>";
                extraClass = " sputil-placeholder";
            }

            return
                $"<div class=\"ms-rtestate-read ms-rte-wpbox{extraClass}\" " +
                $"contenteditable=\"false\" unselectable=\"on\">{metaComment}\r\n" +
                $"  <div class=\"ms-rtestate-notify ms-rtestate-read {zoneKey}\" " +
                $"id=\"div_{zoneKey}\" unselectable=\"on\"></div>\r\n" +
                $"  <div id=\"vid_{zoneKey}\" unselectable=\"on\" style=\"display:none;\"></div>\r\n" +
                $"</div>";
        }

        /// <summary>
        /// Удаляет заглушку с указанным ZoneKey из HTML строки.
        /// </summary>
        private static string RemovePlaceholderFromHtml(string html, string zoneKey)
        {
            // Remove the entire ms-rte-wpbox div containing this ZoneKey
            string pattern =
                @"<div[^>]*ms-rte-wpbox[^>]*>.*?" +
                Regex.Escape(zoneKey) +
                @".*?</div>\s*</div>\s*</div>";

            return Regex.Replace(html, pattern,
                string.Empty, RegexOptions.IgnoreCase | RegexOptions.Singleline);
        }

        /// <summary>
        /// CheckOut с тихим игнорированием если страница уже извлечена.
        /// </summary>
        private async Task SafeCheckOutAsync(ClientContext ctx, Microsoft.SharePoint.Client.File file)
        {
            try
            {
                file.CheckOut();
                await Task.Run(() => ctx.ExecuteQuery());
            }
            catch (ServerException ex) when (
                ex.Message.Contains("already checked out") ||
                ex.Message.Contains("Check out"))
            {
                // Already checked out by us — continue
            }
            catch (ServerException ex) when (
                ex.Message.Contains("checked out for editing") ||
                ex.Message.Contains("is checked out"))
            {
                // Checked out by another user — load file to get checkout info,
                // then take over by discarding their checkout (requires Manage Lists permission)
                System.Diagnostics.Debug.WriteLine(
                    $"[SafeCheckOut] File checked out by another user: {ex.Message}");
                try
                {
                    ctx.Load(file, f => f.CheckOutType);
                    await Task.Run(() => ctx.ExecuteQuery());

                    if (file.CheckOutType != CheckOutType.None)
                    {
                        file.UndoCheckOut();
                        await Task.Run(() => ctx.ExecuteQuery());
                    }

                    // Now check out for ourselves
                    file.CheckOut();
                    await Task.Run(() => ctx.ExecuteQuery());
                }
                catch (Exception innerEx)
                {
                    _log.Caller().Error(innerEx, "ERROR: {ExType} — {Message}", innerEx.GetType().Name, innerEx.Message);
                    throw new InvalidOperationException(
                        $"File is checked out by another user and cannot be taken over automatically.\n" +
                        $"Please ask the user to check it in, or check it in manually via SharePoint UI.\n" +
                        $"Details: {innerEx.Message}", innerEx);
                }
            }
        }

        /// <summary>
        /// CheckIn + Publish.
        /// </summary>
        private async Task CheckInAndPublishAsync(
            ClientContext ctx,
            Microsoft.SharePoint.Client.File file,
            string comment = "")
        {
            file.CheckIn(comment, CheckinType.MajorCheckIn);
            file.Publish(comment);
            await Task.Run(() => ctx.ExecuteQuery());
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  1. GetPageSnapshotAsync
        //     Читает полный снимок страницы с WebParts в визуальном порядке.
        // ═══════════════════════════════════════════════════════════════════════
        // Logged: entry, WebPart count, errors
        public async Task<PageSnapshot> GetPageSnapshotAsync(
            string siteUrl,
            string pageRelativeUrl)
        {
            _log.Debug("GetPageSnapshot: {Page} site={Site}", pageRelativeUrl, siteUrl);
            return await Task.Run(async () =>
            {
                var snapshot = new PageSnapshot
                {
                    PageRelativeUrl = pageRelativeUrl,
                    SnapshotTime    = DateTime.Now
                };

                using var ctx = await GetContextAsync(siteUrl);

                // ── A: Read page list item (Title, Layout, PublishingContent) ──
                var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
                ctx.Load(pageFile, f => f.ListItemAllFields, f => f.Exists);
                await Task.Run(() => ctx.ExecuteQuery());

                // A missing file does not fail here: CSOM returns an object whose
                // ListItemAllFields is never initialized, and the failure surfaces a few
                // lines below as "The property or field has not been initialized" —
                // a message that says nothing about the actual cause.
                if (!pageFile.Exists)
                {
                    _logPage.Caller().Warning("GetPageSnapshot: page not found — {Url}", pageRelativeUrl);
                    throw new System.IO.FileNotFoundException(
                        $"Page not found:\n{pageRelativeUrl}", pageRelativeUrl);
                }
                var fields = pageFile.ListItemAllFields;
                ctx.Load(fields);
                await Task.Run(() => ctx.ExecuteQuery());

                snapshot.PageTitle      = fields["Title"]?.ToString() ?? string.Empty;
                snapshot.PublishingHtml = fields["PublishingPageContent"]?.ToString() ?? string.Empty;

                // Layout is stored as a lookup — server-relative URL is in the URL part
                if (fields.FieldValues.TryGetValue("PublishingPageLayout", out var layoutVal)
                    && layoutVal is FieldUrlValue layoutUrl)
                {
                    snapshot.LayoutRelativeUrl = new Uri(layoutUrl.Url).AbsolutePath;
                }

                // ── B: Read WebParts via LimitedWebPartManager ──
                var wpm = pageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
                ctx.Load(wpm.WebParts, wps => wps.Include(
                    d => d.Id,
                    d => d.ZoneId,
                    d => d.WebPart.Title,
                    d => d.WebPart.Hidden,
					d => d.WebPart.ZoneIndex,
					d => d.WebPart.IsClosed,
                    d => d.WebPart.Properties));
                await Task.Run(() => ctx.ExecuteQuery());
				/*
				foreach (var probe in wpm.WebParts)
                {
                    // Each value is read separately: the server may return one
                    // property and omit another, and an unguarded access would
                    // throw PropertyOrFieldNotInitializedException and abort the copy.
                    string zoneIndexText = SafeRead(() => probe.WebPart.ZoneIndex.ToString());
                    string hiddenText    = SafeRead(() => probe.WebPart.Hidden.ToString());
                    string closedText    = SafeRead(() => probe.WebPart.IsClosed.ToString());

                    System.Diagnostics.Debug.WriteLine(
                        $"[ZONEPROBE] zone='{probe.ZoneId}' index={zoneIndexText} " +
                        $"hidden={hiddenText} closed={closedText} " +
                        $"title='{probe.WebPart.Title}' key={probe.Id:D}");
                }
				*/
                // Build StorageKey → properties dictionary
                var wpByStorageKey = new Dictionary<string, (WebPartDefinition def, Dictionary<string, string> props)>(
                    StringComparer.OrdinalIgnoreCase);

                foreach (var def in wpm.WebParts)
                {
                    var props = def.WebPart.Properties.FieldValues
                        .ToDictionary(kv => kv.Key, kv => kv.Value?.ToString() ?? "");
                    wpByStorageKey[def.Id.ToString("D")] = (def, props);
                }

                // ── C: Fetch rendered HTML to get webpartid / webpartid2 mapping ──
				string renderedHtml;
				try
				{
					renderedHtml = await FetchPageHtmlAsync(siteUrl, pageRelativeUrl);
				}
				catch (HttpRequestException ex) when (ex.Message.Contains("404"))
				{
					// pageFile.Exists was already confirmed true above — the file is real.
					// A 404 specifically on the RENDERED page means it has never been
					// checked in: SharePoint has nothing to serve at this URL outside
					// of edit mode until at least one version is published.
					_logPage.Caller().Warning(
						"GetPageSnapshot: page has no published version — {Url}", pageRelativeUrl);
					throw new InvalidOperationException(
						$"This page has never been checked in.\n\n" +
						$"Save at least one version (Check In) before copying:\n{pageRelativeUrl}");
				}
                var zoneKeyToStorageKey = ParseZoneKeyToStorageKey(renderedHtml);

                // ── DIAGNOSTICS ──────────────────────────────────────────────
                System.Diagnostics.Debug.WriteLine($"[SNAPSHOT] RenderedHtml length: {renderedHtml.Length}");
                System.Diagnostics.Debug.WriteLine($"[SNAPSHOT] HTML first 500 chars:\n{renderedHtml.Substring(0, Math.Min(500, renderedHtml.Length))}");
                System.Diagnostics.Debug.WriteLine($"[SNAPSHOT] zoneKeyToStorageKey count: {zoneKeyToStorageKey.Count}");

                // ── D: Extract visual order from PublishingContent ──
                var zoneKeysInOrder = ParseZoneKeysInOrder(snapshot.PublishingHtml);

                // ── E: Build WebPartSnapshot list in visual order ──
                int position = 1;
                var processedStorageKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var zoneKey in zoneKeysInOrder)
                {
                    // Look up StorageKey for this ZoneKey
                    if (!zoneKeyToStorageKey.TryGetValue(zoneKey, out var storageKey))
                    {
                        // ZoneKey has no matching WebPart object — orphaned placeholder, skip
                        System.Diagnostics.Debug.WriteLine(
                            $"[PageSnapshot] Orphaned placeholder: {zoneKey} — no matching WebPart");
                        continue;
                    }

                    if (!wpByStorageKey.TryGetValue(storageKey, out var wpEntry))
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[PageSnapshot] StorageKey {storageKey} not found in WebPartManager");
                        continue;
                    }

                    // Download ExportXml
                    string exportXml = await GetWebPartExportXmlAsync(siteUrl, pageRelativeUrl, storageKey);

                    snapshot.WebParts.Add(new WebPartSnapshot
                    {
                        StorageKey     = storageKey,
                        ZoneKey        = zoneKey,
                        VisualPosition = position++,
                        ZoneId         = wpEntry.def.ZoneId,
                        // ZoneIndex carries no meaning for WebParts embedded in
                        // PublishingPageContent — their order comes from the position
                        // of the ms-rte-wpbox placeholder in the HTML, not from the zone.
                        // Zeroed deliberately so later code cannot be tempted to use it.
                        ZoneIndex      = 0,
                        IsHidden       = SafeValue(() => wpEntry.def.WebPart.Hidden,   false),
                        IsClosed       = SafeValue(() => wpEntry.def.WebPart.IsClosed, false),
                        Title          = wpEntry.def.WebPart.Title,
                        ExportXml      = exportXml,
                        Properties     = wpEntry.props
                    });

                    processedStorageKeys.Add(storageKey);
                }

                // ── F: Add WebParts from named zones (Header/Right/etc.) ──
                // These WebParts have no ZoneKey in PublishingContent but we still
                // include them with sequential VisualPosition so CreatePageFromSnapshotAsync
                // copies them into PublishingContent on the target.
                foreach (var kv in wpByStorageKey)
                {
                    if (processedStorageKeys.Contains(kv.Key)) continue;

                    string exportXml = await GetWebPartExportXmlAsync(siteUrl, pageRelativeUrl, kv.Key);

					snapshot.WebParts.Add(new WebPartSnapshot
                    {
                        StorageKey     = kv.Key,
                        ZoneKey        = string.Empty,   // was in named zone, no placeholder in PublishingContent
                        VisualPosition = position++,     // sequential — not 0
                        ZoneId         = kv.Value.def.ZoneId,
                        ZoneIndex      = SafeValue(() => kv.Value.def.WebPart.ZoneIndex, 0),
                        IsHidden       = SafeValue(() => kv.Value.def.WebPart.Hidden,   false),
                        IsClosed       = SafeValue(() => kv.Value.def.WebPart.IsClosed, false),
                        Title          = kv.Value.def.WebPart.Title,
                        ExportXml      = exportXml,
                        Properties     = kv.Value.props
                    });
				}

                System.Diagnostics.Debug.WriteLine(
                    $"[PageSnapshot] Done. {snapshot.WebParts.Count} WebParts, " +
                    $"{zoneKeysInOrder.Count} in PublishingContent.");


                // STEP 0.5 VERIFY — confirm the model carries the zone placement data
                foreach (var w in snapshot.WebParts.OrderBy(w => w.ZoneId).ThenBy(w => w.ZoneIndex))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[SNAP] zone='{w.ZoneId}' idx={w.ZoneIndex} zoneKey='{w.ZoneKey}' " +
                        $"vis={w.VisualPosition} hidden={w.IsHidden} closed={w.IsClosed} " +
                        $"xml={(string.IsNullOrEmpty(w.ExportXml) ? "EMPTY" : w.ExportXml.Length + "b")} " +
                        $"title='{w.Title}'");
                }

                return snapshot;
            });
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  11. GetWebPartExportXmlAsync
        //      Скачивает .webpart XML через exportwp.aspx.
        //      Вынесен отдельно — используется во многих методах.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<string> GetWebPartExportXmlAsync(
            string siteUrl,
            string pageRelativeUrl,
            string storageKey)
        {
            string hostRoot  = "https://" + new Uri(siteUrl).Host;
            string pageUrl   = hostRoot + pageRelativeUrl;

            // Normalize both URLs before building the export endpoint
            string normalizedSiteUrl = SPUtil.Infrastructure.SPUsingUtils.NormalizeUrl(siteUrl);
            string normalizedPageUrl = SPUtil.Infrastructure.SPUsingUtils.NormalizeUrl(pageUrl);

            string exportUrl =
                $"{new Uri(normalizedSiteUrl).GetLeftPart(UriPartial.Authority)}" +
                $"{new Uri(normalizedSiteUrl).AbsolutePath.TrimEnd('/')}/" +
                $"_vti_bin/exportwp.aspx" +
                $"?pageurl={Uri.EscapeDataString(normalizedPageUrl)}" +
                $"&guidstring={Uri.EscapeDataString(storageKey)}";

            var handler = new System.Net.Http.HttpClientHandler
            {
                Credentials = GetCredentials(siteUrl)
            };
            using var http = new System.Net.Http.HttpClient(handler);

            var response = await http.GetAsync(exportUrl);
            if (!response.IsSuccessStatusCode)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[ExportXml] HTTP {(int)response.StatusCode} for {storageKey}");
                return string.Empty;
            }

            var bytes  = await response.Content.ReadAsByteArrayAsync();
            var xmlRaw = Encoding.UTF8.GetString(bytes);

            // Strip UTF-8 BOM if present
            return xmlRaw.TrimStart('\uFEFF');
        }

        /*
        // ═══════════════════════════════════════════════════════════════════════
        //  14. CheckOut / CheckIn / Publish как отдельные методы
        // ═══════════════════════════════════════════════════════════════════════
        public async Task CheckOutPageAsync(string siteUrl, string pageRelativeUrl)
        {
            using var ctx = await GetContextAsync(siteUrl);
            var file = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
            ctx.Load(file);
            await Task.Run(() => ctx.ExecuteQuery());
            await SafeCheckOutAsync(ctx, file);
        }

        public async Task CheckInPageAsync(string siteUrl, string pageRelativeUrl, string comment = "")
        {
            using var ctx = await GetContextAsync(siteUrl);
            var file = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
            ctx.Load(file);
            await Task.Run(() => ctx.ExecuteQuery());
            file.CheckIn(comment, CheckinType.MajorCheckIn);
            await Task.Run(() => ctx.ExecuteQuery());
        }

        public async Task PublishPageAsync(string siteUrl, string pageRelativeUrl, string comment = "")
        {
            using var ctx = await GetContextAsync(siteUrl);
            var file = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
            ctx.Load(file);
            await Task.Run(() => ctx.ExecuteQuery());
            file.Publish(comment);
            await Task.Run(() => ctx.ExecuteQuery());
        }
        */

        // ═══════════════════════════════════════════════════════════════════════
        //  5. AddWebPartAsync
        //     Добавляет WebPart на страницу из ExportXml.
        //     Регистрирует в wpz, вставляет заглушку в PublishingContent.
        //     Возвращает StorageKey добавленного WebPart.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<string> AddWebPartAsync(
            string siteUrl,
            string pageRelativeUrl,
            string webPartXml,
            int position = 0)
        {
            return await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(siteUrl);

                var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
                ctx.Load(pageFile, f => f.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());

                await SafeCheckOutAsync(ctx, pageFile);

                bool checkedIn = false;
                try
                {
                    // ── Register WebPart in wpz zone ──
                    var wpm = pageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
                    var imported = wpm.ImportWebPart(webPartXml);
                    var definition = wpm.AddWebPart(imported.WebPart, "wpz", 0);
                    ctx.Load(definition, d => d.Id);
                    await Task.Run(() => ctx.ExecuteQuery());

                    string storageKey = definition.Id.ToString("D");

                    string renderedHtml = await FetchPageHtmlAsync(siteUrl, pageRelativeUrl);
                    var zoneKeyToStorageKey = ParseZoneKeyToStorageKey(renderedHtml);

                    string zoneKey = zoneKeyToStorageKey
                        .FirstOrDefault(kv =>
                            kv.Value.Equals(storageKey, StringComparison.OrdinalIgnoreCase))
                        .Key ?? Guid.NewGuid().ToString("D");

                    ctx.Load(pageFile.ListItemAllFields);
                    await Task.Run(() => ctx.ExecuteQuery());
                    var fields = pageFile.ListItemAllFields;
                    var existing = fields["PublishingPageContent"]?.ToString() ?? "";

                    string placeholder = BuildWpBoxPlaceholder(zoneKey);
                    string newHtml;

                    if (position <= 0 || string.IsNullOrEmpty(existing))
                    {
                        newHtml = existing + "\r\n" + placeholder + "\r\n<p><br/></p>";
                    }
                    else
                    {
                        var zoneKeys = ParseZoneKeysInOrder(existing);

                        if (position > zoneKeys.Count)
                        {
                            newHtml = existing + "\r\n" + placeholder + "\r\n<p><br/></p>";
                        }
                        else
                        {
                            string targetZoneKey = zoneKeys[position - 1];
                            string insertPattern =
                                @"(<div[^>]*ms-rte-wpbox[^>]*>.*?" +
                                Regex.Escape(targetZoneKey) +
                                @".*?</div>\s*</div>\s*</div>)";

                            newHtml = Regex.Replace(
                                existing,
                                insertPattern,
                                placeholder + "\r\n$1",
                                RegexOptions.IgnoreCase | RegexOptions.Singleline);
                        }
                    }

                    fields["PublishingPageContent"] = newHtml;
                    fields.Update();
                    await Task.Run(() => ctx.ExecuteQuery());

                    await CheckInAndPublishAsync(ctx, pageFile, $"Added WebPart {storageKey}");
                    checkedIn = true;

                    System.Diagnostics.Debug.WriteLine(
                        $"[AddWebPart] OK. StorageKey={storageKey}, ZoneKey={zoneKey}");

                    return storageKey;
                }
                finally
                {
                    if (!checkedIn)
                    {
                        try
                        {
                            pageFile.UndoCheckOut();
                            await Task.Run(() => ctx.ExecuteQuery());
                            _logPage.Caller().Warning(
                                "AddWebPartAsync: operation failed — checkout discarded for {Page}",
                                pageRelativeUrl);
                        }
                        catch (Exception cleanupEx)
                        {
                            _logPage.Caller().Error(cleanupEx,
                                "AddWebPartAsync: failed to undo checkout after error — page may remain checked out: {Page}",
                                pageRelativeUrl);
                        }
                    }
                }
            });
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  6. DeleteWebPartAsync
        //     Удаляет WebPart: объект из wpz + заглушку из PublishingContent.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task DeleteWebPartAsync(
            string siteUrl,
            string pageRelativeUrl,
            string storageKey)
        {
            await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(siteUrl);

                var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
                ctx.Load(pageFile, f => f.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());

                await SafeCheckOutAsync(ctx, pageFile);

                bool checkedIn = false;
                try
                {
                    // ── Find ZoneKey before deleting the WebPart object ──
                    string renderedHtml = await FetchPageHtmlAsync(siteUrl, pageRelativeUrl);
                    var zoneKeyToStorageKey = ParseZoneKeyToStorageKey(renderedHtml);

                    string zoneKey = zoneKeyToStorageKey
                        .FirstOrDefault(kv =>
                            kv.Value.Equals(storageKey, StringComparison.OrdinalIgnoreCase))
                        .Key ?? string.Empty;

                    // ── Delete WebPart object from wpz ──
                    var wpm = pageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
                    ctx.Load(wpm.WebParts, wps => wps.Include(d => d.Id));
                    await Task.Run(() => ctx.ExecuteQuery());

                    var definition = wpm.WebParts
                        .FirstOrDefault(d =>
                            d.Id.ToString("D").Equals(storageKey, StringComparison.OrdinalIgnoreCase));

                    if (definition != null)
                    {
                        definition.DeleteWebPart();
                        await Task.Run(() => ctx.ExecuteQuery());
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine(
                            $"[DeleteWebPart] WebPart {storageKey} not found in WebPartManager");
                    }

                    // ── Remove placeholder from PublishingPageContent ──
                    if (!string.IsNullOrEmpty(zoneKey))
                    {
                        ctx.Load(pageFile.ListItemAllFields);
                        await Task.Run(() => ctx.ExecuteQuery());
                        var fields = pageFile.ListItemAllFields;
                        var html = fields["PublishingPageContent"]?.ToString() ?? "";
                        var newHtml = RemovePlaceholderFromHtml(html, zoneKey);

                        fields["PublishingPageContent"] = newHtml;
                        fields.Update();
                        await Task.Run(() => ctx.ExecuteQuery());
                    }

                    await CheckInAndPublishAsync(ctx, pageFile, $"Deleted WebPart {storageKey}");
                    checkedIn = true;

                    System.Diagnostics.Debug.WriteLine(
                        $"[DeleteWebPart] OK. StorageKey={storageKey}, ZoneKey={zoneKey}");
                }
                finally
                {
                    if (!checkedIn)
                    {
                        try
                        {
                            pageFile.UndoCheckOut();
                            await Task.Run(() => ctx.ExecuteQuery());
                            _logPage.Caller().Warning(
                                "DeleteWebPartAsync: operation failed — checkout discarded for {Page}",
                                pageRelativeUrl);
                        }
                        catch (Exception cleanupEx)
                        {
                            _logPage.Caller().Error(cleanupEx,
                                "DeleteWebPartAsync: failed to undo checkout after error — page may remain checked out: {Page}",
                                pageRelativeUrl);
                        }
                    }
                }
            });
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  3. UpdateWebPartAsync
        //     Обновляет Title и/или свойства одного WebPart.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task UpdateWebPartAsync(
            string siteUrl,
            string pageRelativeUrl,
            WebPartUpdateRequest request)
        {
            await UpdateAllWebPartsAsync(siteUrl, pageRelativeUrl,
                new[] { request });
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  4. UpdateAllWebPartsAsync
        //     Применяет список изменений за одну операцию CheckOut/CheckIn.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task UpdateAllWebPartsAsync(
            string siteUrl,
            string pageRelativeUrl,
            IEnumerable<WebPartUpdateRequest> requests)
        {
            await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(siteUrl);

                var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
                ctx.Load(pageFile);
                await Task.Run(() => ctx.ExecuteQuery());

				await SafeCheckOutAsync(ctx, pageFile);

				bool checkedIn = false;
				try
				{
					var wpm = pageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
					ctx.Load(wpm.WebParts, wps => wps.Include(
						d => d.Id,
						d => d.WebPart.Title,
						d => d.WebPart.Properties));
					await Task.Run(() => ctx.ExecuteQuery());

					// Build index for quick lookup
					var defByKey = wpm.WebParts.ToDictionary(
						d => d.Id.ToString("D"),
						d => d,
						StringComparer.OrdinalIgnoreCase);

					bool needsHtmlUpdate = false;
					string currentHtml   = string.Empty;

					foreach (var req in requests)
					{
						if (!defByKey.TryGetValue(req.StorageKey, out var def))
						{
							System.Diagnostics.Debug.WriteLine(
								$"[UpdateWebParts] StorageKey {req.StorageKey} not found");
							continue;
						}

						// Update Title
						if (req.NewTitle != null)
							def.WebPart.Title = req.NewTitle;

						// Update properties — per-property try/catch: some custom WebPart
						// types (e.g. this farm's DynamicForm control) don't expose every
						// base WebPart property via this generic property bag (Height/Width
						// caused "The Height property does not exist" for one such control).
						// Skip just that property and keep applying the rest, rather than
						// failing the whole update — future custom controls may reject
						// different, currently-unknown properties the same way.
						foreach (var kv in req.PropertiesToUpdate)
						{
							try
							{
								// Prefer type-aware conversion: if the target property already
								// has a current value (even a default), its real CLR type is known
								// client-side from that value alone — no server metadata round-trip
								// needed. Converting the incoming string to THAT type before
								// assignment avoids most "cannot be converted to type X" failures
								// (bool/Single/custom enums like c_textAlign, c_formMode). Properties
								// that don't exist on this control at all (Height/Width) have no
								// current value to type-check against, so they fall through to the
								// plain string assignment below and rely on the catch as before.
								object valueToWrite = kv.Value;
								if (def.WebPart.Properties[kv.Key] is object currentValue)
								{
									Type targetType = currentValue.GetType();
									valueToWrite = targetType.IsEnum
										? Enum.Parse(targetType, kv.Value, ignoreCase: true)
										: Convert.ChangeType(kv.Value, targetType);
								}

								def.WebPart.Properties[kv.Key] = valueToWrite;
                                def.SaveWebPartChanges();
                                ctx.ExecuteQuery();
                                _logPage.Caller().Information(
                                    "Written  '{Prop}' in {StorageKey}: ",
                                   kv.Value, kv.Key);

                            }
                            catch (Exception ex)
							{
								_logPage.Caller().Warning(
									"'{Prop}' = '{Key}' : {Message}",
									kv.Key, kv.Value, ex.Message);
							}
						}

						def.SaveWebPartChanges();

						// Reorder if requested
						if (req.NewVisualPosition.HasValue)
							needsHtmlUpdate = true;
					}

					await Task.Run(() => ctx.ExecuteQuery());

					// ── Reorder in HTML if any request asked for it ──
					if (needsHtmlUpdate)
					{
						ctx.Load(pageFile.ListItemAllFields);
						await Task.Run(() => ctx.ExecuteQuery());

						var fields = pageFile.ListItemAllFields;
						currentHtml = fields["PublishingPageContent"]?.ToString() ?? "";

						var zoneKeys = ParseZoneKeysInOrder(currentHtml);

						string renderedHtml       = await FetchPageHtmlAsync(siteUrl, pageRelativeUrl);
						var zoneKeyToStorageKey   = ParseZoneKeyToStorageKey(renderedHtml);
						var storageKeyToZoneKey   = zoneKeyToStorageKey
							.ToDictionary(kv => kv.Value, kv => kv.Key, StringComparer.OrdinalIgnoreCase);

						foreach (var req in requests.Where(r => r.NewVisualPosition.HasValue))
						{
							if (!storageKeyToZoneKey.TryGetValue(req.StorageKey, out var zk)) continue;

							zoneKeys.Remove(zk);
							int insertAt = Math.Min(req.NewVisualPosition!.Value - 1, zoneKeys.Count);
							insertAt     = Math.Max(0, insertAt);
							zoneKeys.Insert(insertAt, zk);
						}

						fields["PublishingPageContent"] = RebuildHtmlInOrder(currentHtml, zoneKeys);
						fields.Update();
						await Task.Run(() => ctx.ExecuteQuery());
					}

					await CheckInAndPublishAsync(ctx, pageFile, "Updated WebParts");
					checkedIn = true;
				}
				finally
				{
					if (!checkedIn)
					{
						try
						{
							pageFile.UndoCheckOut();
							await Task.Run(() => ctx.ExecuteQuery());
							_logPage.Caller().Warning(
								"UpdateAllWebPartsAsync: operation failed — checkout discarded for {Page}",
								pageRelativeUrl);
						}
						catch (Exception cleanupEx)
						{
							_logPage.Caller().Error(cleanupEx,
								"UpdateAllWebPartsAsync: failed to undo checkout after error — page may remain checked out: {Page}",
								pageRelativeUrl);
						}
					}
				}

            });
        }
        // ═══════════════════════════════════════════════════════════════════════
        //  5a. AddWebPartToZoneAsync
        //      Adds a WebPart from its ExportXml into a NAMED LAYOUT ZONE
        //      (Header, RightColumn, CenterLeftColumn_2 …).
        //
        //      Unlike AddWebPartAsync this method deliberately does NOT touch
        //      PublishingPageContent and does NOT fetch the rendered page:
        //      a zone WebPart is rendered by the layout's WebPartZone control and
        //      needs no ms-rte-wpbox placeholder. Writing one would be harmful —
        //      the final rewrite of PublishingPageContent erases such placeholders,
        //      which is exactly how zone WebParts became invisible before.
        //
        //      Returns the StorageKey of the added WebPart.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<string> AddWebPartToZoneAsync(
            string siteUrl,
            string pageRelativeUrl,
            string webPartXml,
            string zoneId,
            int zoneIndex = 0)
        {
            if (string.IsNullOrWhiteSpace(webPartXml))
                throw new ArgumentException("WebPart XML is empty — nothing to add.", nameof(webPartXml));

            if (string.IsNullOrWhiteSpace(zoneId))
                throw new ArgumentException("Zone id is empty.", nameof(zoneId));

            return await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(siteUrl);

                var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
                ctx.Load(pageFile, f => f.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());

				await SafeCheckOutAsync(ctx, pageFile);

				bool checkedIn = false;
				try
				{
					// ── Register the WebPart in the requested layout zone ──
					var wpm        = pageFile.GetLimitedWebPartManager(PersonalizationScope.Shared);
					var imported   = wpm.ImportWebPart(webPartXml);
					var definition = wpm.AddWebPart(imported.WebPart, zoneId, zoneIndex);
					ctx.Load(definition, d => d.Id);
					await Task.Run(() => ctx.ExecuteQuery());

					string storageKey = definition.Id.ToString("D");

					await CheckInAndPublishAsync(ctx, pageFile,
						$"Added WebPart {storageKey} to zone {zoneId}[{zoneIndex}]");
					checkedIn = true;

					_logPage.Information(
						"AddWebPartToZone: {Page} zone={Zone}[{Index}] storageKey={Key}",
						pageRelativeUrl, zoneId, zoneIndex, storageKey);

					System.Diagnostics.Debug.WriteLine(
						$"[AddWebPartToZone] OK. Zone={zoneId}[{zoneIndex}] StorageKey={storageKey}");

					return storageKey;
				}
				finally
				{
					if (!checkedIn)
					{
						try
						{
							pageFile.UndoCheckOut();
							await Task.Run(() => ctx.ExecuteQuery());
							_logPage.Caller().Warning(
								"AddWebPartToZoneAsync: operation failed — checkout discarded for {Page}",
								pageRelativeUrl);
						}
						catch (Exception cleanupEx)
						{
							_logPage.Caller().Error(cleanupEx,
								"AddWebPartToZoneAsync: failed to undo checkout after error — page may remain checked out: {Page}",
								pageRelativeUrl);
						}
					}
				}
            });
        }
        // ═══════════════════════════════════════════════════════════════════════
        //  7. UpdateWebPartPropertyAsync — shortcut для одного свойства
        // ═══════════════════════════════════════════════════════════════════════
        public async Task UpdateWebPartPropertyAsync(
            string siteUrl,
            string pageRelativeUrl,
            string storageKey,
            string propertyName,
            string propertyValue)
        {
            await UpdateWebPartAsync(siteUrl, pageRelativeUrl, new WebPartUpdateRequest
            {
                StorageKey         = storageKey,
                PropertiesToUpdate = new Dictionary<string, string>
                {
                    { propertyName, propertyValue }
                }
            });
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  8. ReorderWebPartsAsync
        //     Переставляет WebParts в PublishingContent по новому порядку.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task ReorderWebPartsAsync(
            string siteUrl,
            string pageRelativeUrl,
            IEnumerable<string> orderedStorageKeys)
        {
            await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(siteUrl);

                var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
                ctx.Load(pageFile, f => f.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());

				await SafeCheckOutAsync(ctx, pageFile);

				bool checkedIn = false;
				try
				{
					// Map StorageKey → ZoneKey via rendered HTML
					string renderedHtml       = await FetchPageHtmlAsync(siteUrl, pageRelativeUrl);
					var zoneKeyToStorageKey   = ParseZoneKeyToStorageKey(renderedHtml);
					var storageKeyToZoneKey   = zoneKeyToStorageKey
						.ToDictionary(kv => kv.Value, kv => kv.Key, StringComparer.OrdinalIgnoreCase);

					// Build new ZoneKey order
					var newZoneKeyOrder = orderedStorageKeys
						.Where(sk => storageKeyToZoneKey.ContainsKey(sk))
						.Select(sk => storageKeyToZoneKey[sk])
						.ToList();

					ctx.Load(pageFile.ListItemAllFields);
					await Task.Run(() => ctx.ExecuteQuery());
					var fields      = pageFile.ListItemAllFields;
					var currentHtml = fields["PublishingPageContent"]?.ToString() ?? "";

					// Add any ZoneKeys not in the new order at the end (safety)
					var existingKeys = ParseZoneKeysInOrder(currentHtml);
					foreach (var zk in existingKeys)
						if (!newZoneKeyOrder.Contains(zk))
							newZoneKeyOrder.Add(zk);

					fields["PublishingPageContent"] = RebuildHtmlInOrder(currentHtml, newZoneKeyOrder);
					fields.Update();
					await Task.Run(() => ctx.ExecuteQuery());

					await CheckInAndPublishAsync(ctx, pageFile, "Reordered WebParts");
					checkedIn = true;
				}
				finally
				{
					if (!checkedIn)
					{
						try
						{
							pageFile.UndoCheckOut();
							await Task.Run(() => ctx.ExecuteQuery());
							_logPage.Caller().Warning(
								"ReorderWebPartsAsync: operation failed — checkout discarded for {Page}",
								pageRelativeUrl);
						}
						catch (Exception cleanupEx)
						{
							_logPage.Caller().Error(cleanupEx,
								"ReorderWebPartsAsync: failed to undo checkout after error — page may remain checked out: {Page}",
								pageRelativeUrl);
						}
					}
				}

            });
        }

        /// <summary>
        /// Перестраивает PublishingContent HTML в новом порядке ZoneKey.
        /// Вырезает каждый wpbox-блок и склеивает заново.
        /// </summary>
        private string RebuildHtmlInOrder(string html, List<string> zoneKeyOrder)
        {
            // Extract each wpbox block keyed by ZoneKey
            var blocks = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var pattern = @"<div[^>]*ms-rte-wpbox[^>]*>.*?(?=" +
                          @"<div[^>]*ms-rte-wpbox|$)";

            var matches = Regex.Matches(html, pattern,
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            foreach (Match m in matches)
            {
                var guidMatch = Regex.Match(m.Value,
                    @"ms-rtestate-read\s+([0-9a-f\-]{36})",
                    RegexOptions.IgnoreCase);
                if (guidMatch.Success)
                    blocks[guidMatch.Groups[1].Value] = m.Value;
            }

            // Strip all wpbox blocks from HTML, keep surrounding text
            string stripped = Regex.Replace(html,
                @"<div[^>]*ms-rte-wpbox[^>]*>.*?</div>\s*</div>\s*</div>",
                "%%WPBOX%%",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            // Remove all placeholders
            stripped = Regex.Replace(stripped, @"%%WPBOX%%", "");
            stripped = stripped.TrimEnd();

            // Append blocks in new order
            var sb = new StringBuilder(stripped);
            foreach (var zk in zoneKeyOrder)
            {
                if (blocks.TryGetValue(zk, out var block))
                {
                    sb.AppendLine();
                    sb.AppendLine(block);
                    sb.AppendLine("<p><br/></p>");
                }
            }

            return sb.ToString();
        }

        /*
        // ═══════════════════════════════════════════════════════════════════════
        //  9. MoveWebPartAsync
        //     Перемещает WebPart между страницами.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task MoveWebPartAsync(
            string siteUrl,
            string sourcePageRelativeUrl,
            string targetPageRelativeUrl,
            string storageKey,
            int targetPosition = 0)
        {
            // Get ExportXml from source
            string xml = await GetWebPartExportXmlAsync(siteUrl, sourcePageRelativeUrl, storageKey);
            if (string.IsNullOrEmpty(xml))
                throw new InvalidOperationException(
                    $"Could not export WebPart {storageKey} from {sourcePageRelativeUrl}");

            // Add to target
            await AddWebPartAsync(siteUrl, targetPageRelativeUrl, xml, targetPosition);

            // Delete from source
            await DeleteWebPartAsync(siteUrl, sourcePageRelativeUrl, storageKey);
        }
        */

        // ═══════════════════════════════════════════════════════════════════════
        //  10. CloneWebPartAsync
        //      Копирует WebPart на ту же страницу с новым StorageKey.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<string> CloneWebPartAsync(
            string siteUrl,
            string pageRelativeUrl,
            string storageKey,
            int targetPosition = 0)
        {
            string xml = await GetWebPartExportXmlAsync(siteUrl, pageRelativeUrl, storageKey);
            if (string.IsNullOrEmpty(xml))
                throw new InvalidOperationException(
                    $"Could not export WebPart {storageKey} from {pageRelativeUrl}");

            return await AddWebPartAsync(siteUrl, pageRelativeUrl, xml, targetPosition);
        }


			// ═══════════════════════════════════════════════════════════════════════
			//  2. CreatePageFromSnapshotAsync
			//     Creates a new publishing page and reproduces every WebPart from the
			//     snapshot. WebParts are split into two groups by ZoneKey: those that
			//     came from PublishingPageContent go back into the content via wpz,
			//     those that came from named layout zones go into their own zones.
			//     Returns a report: what was added, what was skipped, what failed.
			// ═══════════════════════════════════════════════════════════════════════ 
			public async Task<PageCopyReport> CreatePageFromSnapshotAsync(
            string targetSiteUrl,
            string targetPageName,
            PageSnapshot snapshot,
            string subfolderPath = "")   // "" = Pages root; "Dean" or "FacultyAdmin/Sub" = subfolder
			{
            return await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(targetSiteUrl);

                // ── Create Publishing page ──
                var web    = ctx.Web;
                var pubWeb = PublishingWeb.GetPublishingWeb(ctx, web);
                ctx.Load(pubWeb);
                ctx.Load(web, w => w.ServerRelativeUrl);
                await Task.Run(() => ctx.ExecuteQuery());

                // ── Ensure subfolder exists if requested ───────────────────────
                string pagesRoot = web.ServerRelativeUrl.TrimEnd('/') + "/Pages";
                Folder? targetFolder = null;
                if (!string.IsNullOrEmpty(subfolderPath))
                {
                    System.Diagnostics.Debug.WriteLine(
                        $"[CreatePage] Ensuring subfolder: Pages/{subfolderPath}");
                    targetFolder = await EnsureSubfolderAsync(ctx, pagesRoot, subfolderPath);
                }

                // Find the layout in the target Master Page Gallery BY FILE NAME.
                // The snapshot carries the SOURCE server-relative URL
                // (e.g. /home/_catalogs/masterpage/NoMenuPage.aspx), which does not
                // resolve on a different site collection. GetFileByServerRelativeUrl
                // silently returns an object whose ListItemAllFields is null, and the
                // failure only surfaces later inside AddPublishingPage as
                // "Cannot invoke method or retrieve property from null object".
                string layoutFileName = GetLayoutFileName(snapshot.LayoutRelativeUrl);
                if (string.IsNullOrWhiteSpace(layoutFileName))
                    throw new InvalidOperationException(
                        "The source page snapshot contains no page layout — cannot create the target page.");

                var layoutItem = await FindPageLayoutItemAsync(ctx, layoutFileName);
                if (layoutItem == null)
                    throw new InvalidOperationException(
                        $"Page layout '{layoutFileName}' was not found in the Master Page Gallery " +
                        $"of {targetSiteUrl}. Copy the layout to the target site collection first.");

                _logPage.Information(
                    "Layout resolved on target: '{Layout}' (source URL was '{SourceUrl}')",
                    layoutFileName, snapshot.LayoutRelativeUrl);

                var pageInfo = new PublishingPageInformation
                {
                    Name = targetPageName.EndsWith(".aspx")
                                         ? targetPageName
                                         : targetPageName + ".aspx",
                    PageLayoutListItem = layoutItem
                };

                if (targetFolder != null)
                    pageInfo.Folder = targetFolder;

                var newPage = pubWeb.AddPublishingPage(pageInfo);
                ctx.Load(newPage, p => p.ListItem);
                await Task.Run(() => ctx.ExecuteQuery());

                // PublishingPage.Uri does not exist in CSOM —
                // get the server-relative URL from the list item's FileRef field
                ctx.Load(newPage.ListItem, li => li["FileRef"]);
                await Task.Run(() => ctx.ExecuteQuery());

                string newPageRelUrl = newPage.ListItem["FileRef"]?.ToString() ?? string.Empty;

                // Set Title
                var listItem = newPage.ListItem;
                listItem["Title"] = snapshot.PageTitle;
                listItem.Update();
                await Task.Run(() => ctx.ExecuteQuery());

                // AddPublishingPage leaves the page in auto-checkout state.
                // CheckIn immediately to release the lock cleanly.
                var freshFile = ctx.Web.GetFileByServerRelativeUrl(newPageRelUrl);
                ctx.Load(freshFile);
                await Task.Run(() => ctx.ExecuteQuery());
                freshFile.CheckIn("Initial creation", CheckinType.MajorCheckIn);
                await Task.Run(() => ctx.ExecuteQuery());

                System.Diagnostics.Debug.WriteLine(
                    $"[CreatePage] Page created: {newPageRelUrl}");

                // ── Report ────────────────────────────────────────────────────
                var report = new PageCopyReport
                {
                    SourcePageUrl = snapshot.PageRelativeUrl,
                    TargetPageUrl = newPageRelUrl,
                    LayoutName    = layoutFileName
                };

                // ── Split WebParts into two groups ────────────────────────────
                // ZoneKey is the discriminator, not ZoneId: a non-empty ZoneKey means
                // the WebPart was found through its ms-rte-wpbox placeholder inside
                // PublishingPageContent. An empty one means it lives in a named layout
                // zone and has no placeholder at all.
                var inZones = snapshot.WebParts
                    .Where(wp => string.IsNullOrEmpty(wp.ZoneKey))
                    .OrderBy(wp => wp.ZoneId)
                    .ThenBy(wp => wp.ZoneIndex)
                    .ToList();

                var inContent = snapshot.WebParts
                    .Where(wp => !string.IsNullOrEmpty(wp.ZoneKey))
                    .OrderBy(wp => wp.VisualPosition)
                    .ToList();

                _logPage.Information(
                    "CreatePage: {Total} WebPart(s) — {Zones} in layout zones, {Content} in page content",
                    snapshot.WebParts.Count, inZones.Count, inContent.Count);

                // ── Group 1: WebParts in named layout zones ───────────────────
                // Sorted by (ZoneId, ZoneIndex). VisualPosition is meaningless here:
                // it is assigned by the order in which the snapshot happened to walk
                // the WebPart collection, which CSOM does not guarantee.
                foreach (var wp in inZones)
                {
                    var entry = new WebPartCopyEntry
                    {
                        Placement        = WebPartPlacement.LayoutZone,
                        Title            = wp.Title,
                        ZoneId           = wp.ZoneId,
                        Position         = wp.ZoneIndex,
                        SourceStorageKey = wp.StorageKey
                    };

                    if (string.IsNullOrEmpty(wp.ExportXml))
                    {
                        entry.Status = WebPartCopyStatus.Skipped;
                        entry.Reason = "Export returned no XML";
                        _logPage.Caller().Warning(
                            "CreatePage: SKIPPED '{Title}' zone={Zone}[{Index}] key={Key} — empty ExportXml",
                            wp.Title, wp.ZoneId, wp.ZoneIndex, wp.StorageKey);
                    }
                    else
                    {
                        try
                        {
                            entry.TargetStorageKey = await AddWebPartToZoneAsync(
                                targetSiteUrl, newPageRelUrl, wp.ExportXml, wp.ZoneId, wp.ZoneIndex);
                            entry.Status = WebPartCopyStatus.Ok;

                            System.Diagnostics.Debug.WriteLine(
                                $"[CreatePage] Zone WP '{wp.Title}' → {wp.ZoneId}[{wp.ZoneIndex}] " +
                                $"key={entry.TargetStorageKey}");
                        }
                        catch (Exception ex)
                        {
                            // A single WebPart must not abort the whole page:
                            // record the failure and move on to the next one.
                            entry.Status = WebPartCopyStatus.Failed;
                            entry.Reason = ex.Message;
                            _logPage.Caller().Error(ex,
                                "CreatePage: FAILED '{Title}' zone={Zone}[{Index}] key={Key}",
                                wp.Title, wp.ZoneId, wp.ZoneIndex, wp.StorageKey);
                        }
                    }

                    report.Entries.Add(entry);
                }

				// ── Fill missing zone WebParts with a visible warning ──────────────
				// A zone that failed to receive its real WebPart is guaranteed to exist —
				// AddWebPartToZoneAsync would not have been attempted otherwise — so a
				// substitute is always placeable, unlike a text field that may not exist
				// on every page layout.
				var failedZoneEntries = report.Entries
					.Where(e => e.Status == WebPartCopyStatus.Failed && e.Placement == WebPartPlacement.LayoutZone)
					.ToList();

				if (failedZoneEntries.Count > 0)
				{
					int serverMajorVersion = ctx.ServerVersion.Major;

					foreach (var failedEntry in failedZoneEntries)
					{
						string message =
							$"<p>Web Part <b> \"{System.Net.WebUtility.HtmlEncode(failedEntry.Title)}\"</b> not available on this site and cannot be added to this page.<br/>" +
							$"Please install/activate the corresponding feature on this site and copy the page again, " +
							$"or contact IT support: <a href=\"mailto:supportsp@savion.huji.ac.il\">supportsp@savion.huji.ac.il</a></p>"+
							$"<p style=\"font-size:11px;color:#666;\">Technical details: {System.Net.WebUtility.HtmlEncode(failedEntry.Reason)}</p>";

						string warningXml = BuildContentEditorWarningXml(serverMajorVersion, message);

						try
						{
							await AddWebPartToZoneAsync(
								targetSiteUrl, newPageRelUrl, warningXml, failedEntry.ZoneId, failedEntry.Position);
						}
						catch (Exception ex)
						{
							// The warning itself failing to add must not abort page creation —
							// the zone stays empty, but the rest of the page still gets created.
							_logPage.Caller().Error(ex,
								"CreatePage: could not place warning placeholder for '{Title}' in zone={Zone}[{Index}]",
								failedEntry.Title, failedEntry.ZoneId, failedEntry.Position);
						}
					}
				}
                // ── Group 2: WebParts inside PublishingPageContent ────────────
                // Every one of them needs its placeholder GUID remapped: SharePoint
                // assigns a new ZoneKey server-side and it cannot be predicted, so the
                // page has to be rendered after each insertion.
                var oldZoneKeyToNew = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                foreach (var wp in inContent)
                {
                    var entry = new WebPartCopyEntry
                    {
                        Placement        = WebPartPlacement.PageContent,
                        Title            = wp.Title,
                        ZoneId           = "wpz",
						SourceZoneKey    = wp.ZoneKey,
                        Position         = wp.VisualPosition,
                        SourceStorageKey = wp.StorageKey
                    };

                    if (string.IsNullOrEmpty(wp.ExportXml))
                    {
                        entry.Status = WebPartCopyStatus.Skipped;
                        entry.Reason = "Export returned no XML";
                        _logPage.Caller().Warning(
                            "CreatePage: SKIPPED '{Title}' content position {Pos} key={Key} — empty ExportXml",
                            wp.Title, wp.VisualPosition, wp.StorageKey);
                    }
                    else
                    {
                        try
                        {
                            string newStorageKey = await AddWebPartAsync(
                                targetSiteUrl, newPageRelUrl, wp.ExportXml, 0);

                            entry.TargetStorageKey = newStorageKey;
                            entry.Status           = WebPartCopyStatus.Ok;

                            // Resolve the new ZoneKey from the rendered page
                            string renderedHtml     = await FetchPageHtmlAsync(targetSiteUrl, newPageRelUrl);
                            var zoneKeyToStorageKey = ParseZoneKeyToStorageKey(renderedHtml);
                            string newZoneKey       = zoneKeyToStorageKey
                                .FirstOrDefault(kv =>
                                    kv.Value.Equals(newStorageKey, StringComparison.OrdinalIgnoreCase))
                                .Key ?? string.Empty;

                            if (!string.IsNullOrEmpty(newZoneKey))
                            {
                                oldZoneKeyToNew[wp.ZoneKey] = newZoneKey;
                            }
                            else
                            {
                                // The WebPart exists on the page but its placeholder
                                // cannot be remapped — it will not be visible.
                                _logPage.Caller().Warning(
                                    "CreatePage: '{Title}' added but its new ZoneKey was not found in " +
                                    "the rendered page — placeholder will not be remapped",
                                    wp.Title);
                            }

                            System.Diagnostics.Debug.WriteLine(
                                $"[CreatePage] Content WP '{wp.Title}' added. " +
                                $"OldZK={wp.ZoneKey} NewZK={newZoneKey}");
                        }
                        catch (Exception ex)
                        {
                            entry.Status = WebPartCopyStatus.Failed;
                            entry.Reason = ex.Message;
                            _logPage.Caller().Error(ex,
                                "CreatePage: FAILED '{Title}' content position {Pos} key={Key}",
                                wp.Title, wp.VisualPosition, wp.StorageKey);
                        }
                    }

                    report.Entries.Add(entry);
                }


            // ── Rebuild PublishingHtml with new ZoneKeys ──────────────────
            // Replace all old ZoneKey GUIDs in the snapshot HTML with new ones.
            string newHtml = snapshot.PublishingHtml;
                foreach (var kv in oldZoneKeyToNew)
                    newHtml = newHtml.Replace(kv.Key, kv.Value, StringComparison.OrdinalIgnoreCase);

                // Placeholders of WebParts that failed to import keep their old GUID and
                // render as empty boxes — swap them for a visible warning instead.
                newHtml = ReplaceFailedWebPartPlaceholders(
                    newHtml,
                    report.Entries.Where(e =>
                        e.Status == WebPartCopyStatus.Failed &&
                        e.Placement == WebPartPlacement.PageContent));

				using var ctx2 = await GetContextAsync(targetSiteUrl);

				var pageFile2 = ctx2.Web.GetFileByServerRelativeUrl(newPageRelUrl);
				ctx2.Load(pageFile2, f => f.ListItemAllFields);
				await Task.Run(() => ctx2.ExecuteQuery());

				// Page was checked in above — now check it out cleanly for editing
				await SafeCheckOutAsync(ctx2, pageFile2);

				bool checkedIn = false;
				try
				{
					ctx2.Load(pageFile2.ListItemAllFields);
					await Task.Run(() => ctx2.ExecuteQuery());

					var fields2 = pageFile2.ListItemAllFields;
					fields2["PublishingPageContent"] = newHtml;
					fields2.Update();
					await Task.Run(() => ctx2.ExecuteQuery());

					await CheckInAndPublishAsync(ctx2, pageFile2,
						$"Created from snapshot of {snapshot.PageRelativeUrl}");
					checkedIn = true;
				}
				finally
				{
					if (!checkedIn)
					{
						try
						{
							pageFile2.UndoCheckOut();
							await Task.Run(() => ctx2.ExecuteQuery());
							_logPage.Caller().Warning(
								"CreatePageFromSnapshotAsync: final content write failed — checkout discarded for {Page}",
								newPageRelUrl);
						}
						catch (Exception cleanupEx)
						{
							_logPage.Caller().Error(cleanupEx,
								"CreatePageFromSnapshotAsync: failed to undo checkout after error — page may remain checked out: {Page}",
								newPageRelUrl);
						}
					}
				}

				_logPage.Information("CreatePage: done. {Summary}", report.Summary);

				System.Diagnostics.Debug.WriteLine(
					$"[CreatePage] Done. {report.Summary}");

				return report;

            });
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  12. RestorePageFromSnapshotAsync
        //      Применяет снимок к существующей странице (replace all WebParts).
        // ═══════════════════════════════════════════════════════════════════════
        public async Task RestorePageFromSnapshotAsync(
            string siteUrl,
            string pageRelativeUrl,
            PageSnapshot snapshot)
        {
            // Step 1: Delete all current WebParts
            var currentSnapshot = await GetPageSnapshotAsync(siteUrl, pageRelativeUrl);
            foreach (var wp in currentSnapshot.WebParts)
                await DeleteWebPartAsync(siteUrl, pageRelativeUrl, wp.StorageKey);

            // Step 2: Add all WebParts from snapshot in visual order
            foreach (var wp in snapshot.WebParts
                .Where(w => w.VisualPosition > 0)
                .OrderBy(w => w.VisualPosition))
            {
                if (!string.IsNullOrEmpty(wp.ExportXml))
                    await AddWebPartAsync(siteUrl, pageRelativeUrl, wp.ExportXml, 0);
            }
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  13. ComparePageSnapshotsAsync
        //      Текстовый diff двух снимков страниц.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<string> ComparePageSnapshotsAsync(
            PageSnapshot source,
            PageSnapshot target)
        {
            return await Task.Run(() =>
            {
                var sb = new StringBuilder();
                sb.AppendLine($"=== Page Snapshot Comparison ===");
                sb.AppendLine($"Source : {source.PageRelativeUrl}  ({source.SnapshotTime:dd.MM.yyyy HH:mm})");
                sb.AppendLine($"Target : {target.PageRelativeUrl}  ({target.SnapshotTime:dd.MM.yyyy HH:mm})");
                sb.AppendLine(new string('═', 60));

                // Index by Title (Title is the most stable identifier across sites)
                var sourceByTitle = source.WebParts.ToDictionary(w => w.Title, StringComparer.OrdinalIgnoreCase);
                var targetByTitle = target.WebParts.ToDictionary(w => w.Title, StringComparer.OrdinalIgnoreCase);

                // Added in target
                var added = targetByTitle.Keys.Except(sourceByTitle.Keys, StringComparer.OrdinalIgnoreCase).ToList();
                if (added.Any())
                {
                    sb.AppendLine("\n➕ Added in target:");
                    foreach (var t in added)
                        sb.AppendLine($"   + {t}  (pos {targetByTitle[t].VisualPosition})");
                }

                // Removed from source
                var removed = sourceByTitle.Keys.Except(targetByTitle.Keys, StringComparer.OrdinalIgnoreCase).ToList();
                if (removed.Any())
                {
                    sb.AppendLine("\n➖ Removed (present in source, missing in target):");
                    foreach (var r in removed)
                        sb.AppendLine($"   - {r}  (pos {sourceByTitle[r].VisualPosition})");
                }

                // Changed properties
                var common = sourceByTitle.Keys.Intersect(targetByTitle.Keys, StringComparer.OrdinalIgnoreCase).ToList();
                foreach (var title in common)
                {
                    var srcWp = sourceByTitle[title];
                    var tgtWp = targetByTitle[title];
                    var diffs = new List<string>();

                    // Position change
                    if (srcWp.VisualPosition != tgtWp.VisualPosition)
                        diffs.Add($"Position: {srcWp.VisualPosition} → {tgtWp.VisualPosition}");

                    // Property changes
                    var allKeys = srcWp.Properties.Keys.Union(tgtWp.Properties.Keys, StringComparer.OrdinalIgnoreCase);
                    foreach (var key in allKeys)
                    {
                        srcWp.Properties.TryGetValue(key, out var sv);
                        tgtWp.Properties.TryGetValue(key, out var tv);
                        if (!string.Equals(sv ?? "", tv ?? "", StringComparison.Ordinal))
                            diffs.Add($"{key}: \"{sv}\" → \"{tv}\"");
                    }

                    if (diffs.Any())
                    {
                        sb.AppendLine($"\n✏ Changed: {title}");
                        foreach (var d in diffs)
                            sb.AppendLine($"   {d}");
                    }
                }

                if (!added.Any() && !removed.Any() && !common.Any(t =>
                {
                    var s = sourceByTitle[t]; var tg = targetByTitle[t];
                    return s.VisualPosition != tg.VisualPosition ||
                           s.Properties.Any(kv => tg.Properties.TryGetValue(kv.Key, out var v) && v != kv.Value);
                }))
                {
                    sb.AppendLine("\n✔ No differences found.");
                }

                return sb.ToString();
            });
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  15. GetAllPagesSnapshotsAsync
        //      Читает снимки всех страниц из библиотеки Pages.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<List<PageSnapshot>> GetAllPagesSnapshotsAsync(string siteUrl)
        {
            var result = new List<PageSnapshot>();

            // Get all pages from the Pages library
            using var ctx = await GetContextAsync(siteUrl);
            var web   = ctx.Web;
            ctx.Load(web, w => w.ServerRelativeUrl);
            await Task.Run(() => ctx.ExecuteQuery());

            string pagesLibUrl = web.ServerRelativeUrl.TrimEnd('/') + "/Pages";
            var pagesList      = web.GetList(pagesLibUrl);

            var caml = new CamlQuery
            {
                ViewXml = @"<View Scope='RecursiveAll'>
                                <Query>
                                    <Where>
                                        <Eq>
                                            <FieldRef Name='FSObjType'/>
                                            <Value Type='Integer'>0</Value>
                                        </Eq>
                                    </Where>
                                </Query>
                            </View>"
            };

            var items = pagesList.GetItems(caml);
            ctx.Load(items, ii => ii.Include(
                i => i["FileRef"],
                i => i["FileLeafRef"],
				i => i["_ModerationStatus"],
                i => i["Title"]));
            await Task.Run(() => ctx.ExecuteQuery());

            foreach (var item in items)
            {
                string fileRef = item["FileRef"]?.ToString() ?? "";
                if (!fileRef.EndsWith(".aspx", StringComparison.OrdinalIgnoreCase))
                    continue;

                try
                {
					string moderationStatus = item["_ModerationStatus"]?.ToString() ?? "";
					string status = moderationStatus switch
					{
						"0" => "Published",
						"1" => "Rejected",
						"2" => "Pending Approval",
						"3" => "Draft",
						_   => "Unknown"
					};					
                    var snap = await GetPageSnapshotAsync(siteUrl, fileRef);
                    result.Add(snap);
                    System.Diagnostics.Debug.WriteLine(
                        $"[GetAllPages] Snapped: {fileRef} ({snap.WebParts.Count} WPs)");
                }
                catch (Exception ex)
                {
                    _log.Caller().Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                    System.Diagnostics.Debug.WriteLine(
                        $"[GetAllPages] Error on {fileRef}: {ex.Message}");
                }
            }

            return result;
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  ParsePlaceholderMeta
        //  Reads PublishingPageContent and extracts all <!--SPUTIL:{...}--> comments.
        //  Returns list of WebPartPlaceholderMeta with TargetZoneKey filled in.
        // ═══════════════════════════════════════════════════════════════════════
        private async Task<List<WebPartPlaceholderMeta>> ParsePlaceholderMetaAsync(
            string siteUrl,
            string pageRelativeUrl)
        {
            var result = new List<WebPartPlaceholderMeta>();

            using var ctx = await GetContextAsync(siteUrl);
            var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
            ctx.Load(pageFile, f => f.ListItemAllFields);
            await Task.Run(() => ctx.ExecuteQuery());
            ctx.Load(pageFile.ListItemAllFields);
            await Task.Run(() => ctx.ExecuteQuery());

            string html = pageFile.ListItemAllFields["PublishingPageContent"]?.ToString() ?? "";
            if (string.IsNullOrEmpty(html)) return result;

            // Match each SPUTIL comment with the ZoneKey of its containing wpbox div
            // Pattern: find ms-rte-wpbox div, extract ZoneKey from ms-rtestate-read class,
            // and look for <!--SPUTIL:{...}--> inside the same div block.
            var divMatches = Regex.Matches(
                html,
                @"<div[^>]*ms-rte-wpbox[^>]*>(.*?)</div>\s*</div>\s*</div>",
                RegexOptions.IgnoreCase | RegexOptions.Singleline);

            foreach (Match div in divMatches)
            {
                string block = div.Value;

                // Find SPUTIL comment
                var commentMatch = Regex.Match(block,
                    @"<!--SPUTIL:(\{[^}]+\})-->",
                    RegexOptions.IgnoreCase);
                if (!commentMatch.Success) continue;

                string json = commentMatch.Groups[1].Value;

                // Find ZoneKey from ms-rtestate-read class
                var zoneKeyMatch = Regex.Match(block,
                    @"ms-rtestate-read\s+([0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12})",
                    RegexOptions.IgnoreCase);

                string targetZoneKey = zoneKeyMatch.Success
                    ? zoneKeyMatch.Groups[1].Value.ToLower()
                    : string.Empty;

                try
                {
                    // Simple JSON parse without external dependencies
                    var meta = new WebPartPlaceholderMeta
                    {
                        TargetZoneKey = targetZoneKey,
                        StorageKey    = ExtractJsonString(json, "storageKey"),
                        Title         = ExtractJsonString(json, "title"),
                        Position      = int.TryParse(ExtractJsonString(json, "position"), out var pos) ? pos : 0,
                        ZoneId        = ExtractJsonString(json, "zoneId"),
                        SiteUrl       = ExtractJsonString(json, "siteUrl"),
                        PageUrl       = ExtractJsonString(json, "pageUrl")
                    };
                    result.Add(meta);
                }
                catch (Exception ex)
                {
                    _log.Caller().Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                    System.Diagnostics.Debug.WriteLine($"[ParsePlaceholder] Failed to parse: {json} | {ex.Message}");
                }
            }

            return result;
        }

        /// <summary>Extracts a string value from a simple JSON object by key name.</summary>
        private static string ExtractJsonString(string json, string key)
        {
            var m = Regex.Match(json,
                $@"""{Regex.Escape(key)}""\s*:\s*""([^""]*)""",
                RegexOptions.IgnoreCase);
            if (m.Success) return m.Groups[1].Value;

            // Try numeric value (for position)
            var mn = Regex.Match(json,
                $@"""{Regex.Escape(key)}""\s*:\s*(\d+)",
                RegexOptions.IgnoreCase);
            return mn.Success ? mn.Groups[1].Value : string.Empty;
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  PageHasPlaceholdersAsync
        //  Returns true if the page contains any SPUTIL placeholder comments.
        //  Used to enable/disable the Sync Properties button in the toolbar.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<bool> PageHasPlaceholdersAsync(string siteUrl, string pageRelativeUrl)
        {
            using var ctx = await GetContextAsync(siteUrl);
            var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
            ctx.Load(pageFile, f => f.ListItemAllFields);
            await Task.Run(() => ctx.ExecuteQuery());
            ctx.Load(pageFile.ListItemAllFields);
            await Task.Run(() => ctx.ExecuteQuery());

            string html = pageFile.ListItemAllFields["PublishingPageContent"]?.ToString() ?? "";
            return html.Contains("<!--SPUTIL:", StringComparison.OrdinalIgnoreCase);
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  ComparePageSnapshotsStructured
        //  Returns structured PageCompareResult.
        //  Matching by Title — position is used as tiebreaker for duplicates.
        // ═══════════════════════════════════════════════════════════════════════
        public Task<PageCompareResult> ComparePageSnapshotsStructured(
            PageSnapshot source,
            PageSnapshot target,
            string sourceSiteUrl = "",
            string targetSiteUrl = "")
        {
            return Task.Run(() =>
            {
                var result = new PageCompareResult
                {
                    SourceUrl          = source.PageRelativeUrl,
                    TargetUrl          = target.PageRelativeUrl,
                    SourceSiteUrl      = sourceSiteUrl,
                    TargetSiteUrl      = targetSiteUrl,
                    LayoutMatches      = string.Equals(
                        source.LayoutRelativeUrl, target.LayoutRelativeUrl,
                        StringComparison.OrdinalIgnoreCase),
                    SourceLayout       = source.LayoutRelativeUrl,
                    TargetLayout       = target.LayoutRelativeUrl,
                    SourceWebPartCount = source.WebParts.Count,
                    TargetWebPartCount = target.WebParts.Count
                };

                // Properties to skip — always differ between sites
                var skipProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
                {
                    "AllowClose","AllowConnect","AllowEdit","AllowHide",
                    "AllowMinimize","AllowZoneChange","AuthorizationFilter",
                    "CatalogIconImageUrl","ChromeState","ChromeType",
                    "Direction","ExportMode","HelpMode","HelpUrl",
                    "Hidden","ImportErrorMessage","TitleIconImageUrl","TitleUrl"
                };

                // Build target lookup: Title → list of WebPartSnapshot (handle duplicates)
                var targetByTitle = new Dictionary<string, List<WebPartSnapshot>>(StringComparer.OrdinalIgnoreCase);
                foreach (var tw in target.WebParts.OrderBy(w => w.VisualPosition))
                {
                    if (!targetByTitle.ContainsKey(tw.Title))
                        targetByTitle[tw.Title] = new List<WebPartSnapshot>();
                    targetByTitle[tw.Title].Add(tw);
                }

                var matchedTargetKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                // ── Process each source WebPart ───────────────────────────────
                foreach (var sw in source.WebParts.OrderBy(w => w.VisualPosition))
                {
                    if (!targetByTitle.TryGetValue(sw.Title, out var candidates))
                    {
                        // Not on target → Removed
                        result.Diffs.Add(new WebPartDiff
                        {
                            Kind             = WebPartDiffKind.Removed,
                            Title            = sw.Title,
                            SourcePosition   = sw.VisualPosition,
                            TargetPosition   = 0,
                            SourceStorageKey = sw.StorageKey,
                            SourceZoneId     = sw.ZoneId
                        });
                        continue;
                    }

                    // Match by closest position (handles duplicates)
                    var tw = candidates
                        .Where(c => !matchedTargetKeys.Contains(c.StorageKey))
                        .OrderBy(c => Math.Abs(c.VisualPosition - sw.VisualPosition))
                        .FirstOrDefault();

                    if (tw == null)
                    {
                        // All candidates already matched → treat as Removed
                        result.Diffs.Add(new WebPartDiff
                        {
                            Kind             = WebPartDiffKind.Removed,
                            Title            = sw.Title,
                            SourcePosition   = sw.VisualPosition,
                            TargetPosition   = 0,
                            SourceStorageKey = sw.StorageKey,
                            SourceZoneId     = sw.ZoneId
                        });
                        continue;
                    }

                    matchedTargetKeys.Add(tw.StorageKey);

                    // Compare properties
                    var propDiffs = new List<PropertyDiff>();
                    var allKeys = sw.Properties.Keys
                        .Union(tw.Properties.Keys, StringComparer.OrdinalIgnoreCase)
                        .Where(k => !skipProps.Contains(k))
                        .OrderBy(k => k);

                    foreach (var key in allKeys)
                    {
                        sw.Properties.TryGetValue(key, out var sv); sv ??= "";
                        tw.Properties.TryGetValue(key, out var tv); tv ??= "";
                        if (!string.Equals(sv, tv, StringComparison.Ordinal))
                            propDiffs.Add(new PropertyDiff
                            {
                                PropertyName = key,
                                SourceValue  = sv,
                                TargetValue  = tv
                            });
                    }

                    bool posChanged = sw.VisualPosition != tw.VisualPosition;

                    result.Diffs.Add(new WebPartDiff
                    {
                        Kind             = (propDiffs.Any() || posChanged)
                                           ? WebPartDiffKind.Modified
                                           : WebPartDiffKind.Identical,
                        Title            = sw.Title,
                        SourcePosition   = sw.VisualPosition,
                        TargetPosition   = tw.VisualPosition,
                        SourceStorageKey = sw.StorageKey,
                        SourceZoneId     = sw.ZoneId,
                        PropertyDiffs    = propDiffs
                    });
                }

                // ── WebParts on target not matched to any source → Added ───────
                foreach (var tw in target.WebParts)
                {
                    if (!matchedTargetKeys.Contains(tw.StorageKey))
                    {
                        result.Diffs.Add(new WebPartDiff
                        {
                            Kind           = WebPartDiffKind.Added,
                            Title          = tw.Title,
                            SourcePosition = 0,
                            TargetPosition = tw.VisualPosition
                        });
                    }
                }

                // Sort: Removed → Added → Modified → Identical, then by position
                result.Diffs = result.Diffs
                    .OrderBy(d => d.Kind switch
                    {
                        WebPartDiffKind.Removed  => 0,
                        WebPartDiffKind.Added    => 1,
                        WebPartDiffKind.Modified => 2,
                        _                        => 3
                    })
                    .ThenBy(d => d.SourcePosition > 0 ? d.SourcePosition : d.TargetPosition)
                    .ToList();

                return result;
            });
        }

        // ═══════════════════════════════════════════════════════════════════════
        //  FormatCopyReport
        //  Renders a PageCopyReport as plain text for UniversalPreviewWindow.
        //  Shown to the user only when something went wrong; the log always keeps
        //  the full picture.
        // ═══════════════════════════════════════════════════════════════════════
        public string FormatCopyReport(PageCopyReport r)
        {
            var sb = new System.Text.StringBuilder();

            sb.AppendLine("=== Page Copy Report ===");
            sb.AppendLine($"Source : {r.SourcePageUrl}");
            sb.AppendLine($"Target : {r.TargetPageUrl}");
            sb.AppendLine($"Layout : {r.LayoutName}");
            sb.AppendLine($"Time   : {r.CopyTime:dd.MM.yyyy HH:mm:ss}");
            sb.AppendLine(new string('═', 70));
            sb.AppendLine();

			sb.AppendLine($"WebParts: {r.TotalCount} total — " +
              $"{r.OkCount} copied, {r.TotalCount - r.OkCount} failed");
			  
			  sb.AppendLine();

            // Same order as the WebParts grid: page content first, then layout zones
            // grouped by zone. The order in which the copy processed them is different
            // and would only confuse a side-by-side comparison with the grid.
            var ordered = r.Entries
                .OrderBy(e => e.Placement == WebPartPlacement.PageContent ? 0 : 1)
                .ThenBy(e => e.Placement == WebPartPlacement.PageContent ? e.Position : 0)
                .ThenBy(e => e.ZoneId, StringComparer.OrdinalIgnoreCase)
                .ThenBy(e => e.Position)
                .ToList();

            sb.AppendLine("Status   Zone                  Idx  Title");
            sb.AppendLine("───────  ────────────────────  ───  ──────────────────────────────");

            foreach (var e in ordered)
            {
				string status = e.Status == WebPartCopyStatus.Ok ? "COPIED" : "FAILED";

                sb.AppendLine($"{status,-7}  {Trim(e.ZoneId, 20),-20}  {e.Position,3}  {e.Title}");
            }

            // ── Problems ──────────────────────────────────────────────────────
            var problems = ordered
                .Where(e => e.Status != WebPartCopyStatus.Ok)
                .ToList();

            if (problems.Count > 0)
            {
                sb.AppendLine(new string('═', 70));
                sb.AppendLine();
                sb.AppendLine("Problems:");

                foreach (var e in problems)
                {
					sb.AppendLine();
					sb.AppendLine($"  FAILED  {e.Title}  ({e.ZoneId}[{e.Position}])");
					sb.AppendLine($"           {e.Reason}");
                    // The source key locates the WebPart on the source page via
                    // ?contents=1 or exportwp.aspx — the only reliable handle when
                    // several WebParts share a title.
                    sb.AppendLine($"           Source key: {e.SourceStorageKey}");

                    if (!string.IsNullOrEmpty(e.TargetStorageKey))
                        sb.AppendLine($"           Target key: {e.TargetStorageKey}");
                }
            }

            sb.AppendLine();
            sb.AppendLine(new string('═', 70));

            return sb.ToString();

            // Zone names can exceed the column width (CenterLeftColumn_2 is 18 chars,
            // custom layouts may go further) — cut rather than break the table.
            static string Trim(string s, int max) =>
                string.IsNullOrEmpty(s) ? string.Empty
                : s.Length <= max ? s : s.Substring(0, max - 1) + "…";
        }
        // ═══════════════════════════════════════════════════════════════════════
        //  FormatCompareResult
        //  Formats PageCompareResult into readable text for UniversalPreviewWindow.
        // ═══════════════════════════════════════════════════════════════════════
        public string FormatCompareResult(PageCompareResult r)
        {
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== Page WebPart Comparison ===");
            sb.AppendLine($"Source : {r.SourceSiteUrl}{r.SourceUrl}");
            sb.AppendLine($"Target : {r.TargetSiteUrl}{r.TargetUrl}");
            sb.AppendLine($"Generated: {System.DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            sb.AppendLine(new string('═', 70));

            sb.AppendLine();
            sb.AppendLine("── Summary ──────────────────────────────────────────────────────────");
            sb.AppendLine($"  Layout       : {(r.LayoutMatches ? "✔ Match" : $"✘ Differs  ({r.SourceLayout}  vs  {r.TargetLayout})")}");
            sb.AppendLine($"  Source WPs   : {r.SourceWebPartCount}");
            sb.AppendLine($"  Target WPs   : {r.TargetWebPartCount}");
            sb.AppendLine();
            sb.AppendLine($"  ✔ Identical  : {r.IdenticalCount}");
            sb.AppendLine($"  ✏ Modified   : {r.ModifiedCount}");
            sb.AppendLine($"  ➕ Added      : {r.AddedCount}");
            sb.AppendLine($"  ➖ Removed    : {r.RemovedCount}");

            if (r.IsIdentical)
            {
                sb.AppendLine();
                sb.AppendLine("✔ Pages are IDENTICAL.");
                return sb.ToString();
            }

            sb.AppendLine();
            sb.AppendLine("── Details ──────────────────────────────────────────────────────────");

            foreach (var d in r.Diffs)
            {
                sb.AppendLine();
                string icon = d.Kind switch
                {
                    WebPartDiffKind.Identical => "✔",
                    WebPartDiffKind.Modified  => "✏",
                    WebPartDiffKind.Added     => "➕",
                    WebPartDiffKind.Removed   => "➖",
                    _                         => "?"
                };

                string posInfo = d.Kind switch
                {
                    WebPartDiffKind.Removed  => $"source pos {d.SourcePosition}  →  MISSING on target",
                    WebPartDiffKind.Added    => $"MISSING on source  →  target pos {d.TargetPosition}",
                    WebPartDiffKind.Modified => d.SourcePosition != d.TargetPosition
                                               ? $"pos {d.SourcePosition} (src) → {d.TargetPosition} (tgt)  ⚠ position changed"
                                               : $"pos {d.SourcePosition}",
                    _                        => $"pos {d.SourcePosition}"
                };

                sb.AppendLine($"{icon} [{d.Kind}]  {d.Title}");
                sb.AppendLine($"   {posInfo}");

                if (d.PropertyDiffs.Any())
                {
                    sb.AppendLine("   ── Changed properties ──────────────────────────────────────");
                    foreach (var pd in d.PropertyDiffs)
                    {
                        string sv = string.IsNullOrEmpty(pd.SourceValue) ? "(empty)" : pd.SourceValue;
                        string tv = string.IsNullOrEmpty(pd.TargetValue) ? "(empty)" : pd.TargetValue;
                        sb.AppendLine($"   {pd.PropertyName,-32}");
                        sb.AppendLine($"     src: {sv}");
                        sb.AppendLine($"     tgt: {tv}");
                    }
                }

                sb.AppendLine(new string('─', 70));
            }

            return sb.ToString();
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  InsertPlaceholdersAsync
        //  For each Removed WebPart in compareResult — inserts a placeholder div
        //  with <!--SPUTIL:{...}--> metadata into target PublishingPageContent.
        //  Placeholders are inserted at the same visual position as on source.
        // ═══════════════════════════════════════════════════════════════════════
        // Logged: placeholders inserted count and errors
        public async Task InsertPlaceholdersAsync(
            string targetSiteUrl,
            string targetPageRelativeUrl,
            PageCompareResult compareResult)
        {
            var removed = compareResult.RemovedWebParts.ToList();
            if (!removed.Any()) return;

            await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(targetSiteUrl);
                var pageFile = ctx.Web.GetFileByServerRelativeUrl(targetPageRelativeUrl);
                ctx.Load(pageFile, f => f.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());

				await SafeCheckOutAsync(ctx, pageFile);

				bool checkedIn = false;
				try
				{
					ctx.Load(pageFile.ListItemAllFields);
					await Task.Run(() => ctx.ExecuteQuery());

					string html = pageFile.ListItemAllFields["PublishingPageContent"]?.ToString() ?? "";

					// Insert placeholders in reverse position order so earlier insertions
					// don't shift the indices of later ones
					foreach (var diff in removed.OrderByDescending(d => d.SourcePosition))
					{
						var meta = new WebPartPlaceholderMeta
						{
							StorageKey = diff.SourceStorageKey,
							Title      = diff.Title,
							Position   = diff.SourcePosition,
							ZoneId     = diff.SourceZoneId,
							SiteUrl    = compareResult.SourceSiteUrl,
							PageUrl    = compareResult.SourceUrl
						};

						string placeholderZoneKey = Guid.NewGuid().ToString("D");
						string placeholder = BuildWpBoxPlaceholder(placeholderZoneKey, meta);

						var existingZoneKeys = ParseZoneKeysInOrder(html);
						int insertAfterIndex = diff.SourcePosition - 1;

						if (existingZoneKeys.Count == 0 || insertAfterIndex <= 0)
						{
							html = placeholder + "\r\n<p><br/></p>\r\n" + html;
						}
						else if (insertAfterIndex >= existingZoneKeys.Count)
						{
							html = html + "\r\n" + placeholder + "\r\n<p><br/></p>";
						}
						else
						{
							string anchorKey = existingZoneKeys[insertAfterIndex - 1];
							string insertPattern =
								@"(<div[^>]*ms-rte-wpbox[^>]*>.*?" +
								Regex.Escape(anchorKey) +
								@".*?</div>\s*</div>\s*</div>)";

							html = Regex.Replace(
								html,
								insertPattern,
								"$1\r\n" + placeholder,
								RegexOptions.IgnoreCase | RegexOptions.Singleline);
						}
					}

					var fields = pageFile.ListItemAllFields;
					fields["PublishingPageContent"] = html;
					fields.Update();
					await Task.Run(() => ctx.ExecuteQuery());

					await CheckInAndPublishAsync(ctx, pageFile,
						$"Inserted {removed.Count} WebPart placeholder(s)");
					checkedIn = true;

					System.Diagnostics.Debug.WriteLine(
						$"[InsertPlaceholders] Done. {removed.Count} placeholders inserted.");
				}
				finally
				{
					if (!checkedIn)
					{
						try
						{
							pageFile.UndoCheckOut();
							await Task.Run(() => ctx.ExecuteQuery());
							_logPage.Caller().Warning(
								"InsertPlaceholdersAsync: operation failed — checkout discarded for {Page}",
								targetPageRelativeUrl);
						}
						catch (Exception cleanupEx)
						{
							_logPage.Caller().Error(cleanupEx,
								"InsertPlaceholdersAsync: failed to undo checkout after error — page may remain checked out: {Page}",
								targetPageRelativeUrl);
						}
					}
				}
            });
        }


        // ═══════════════════════════════════════════════════════════════════════
        //  SyncPropertiesAsync
        //  Finds all SPUTIL placeholders on the target page,
        //  reads WebPart settings from source via exportwp.aspx,
        //  applies them to matching WebParts on target,
        //  then removes the placeholder.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<SyncResult> SyncPropertiesAsync(
            string targetSiteUrl,
            string targetPageRelativeUrl)
        {
            var syncResult = new SyncResult();

            // ── Step 1: Parse all placeholders from target PublishingContent ──
            var placeholders = await ParsePlaceholderMetaAsync(targetSiteUrl, targetPageRelativeUrl);
            if (!placeholders.Any())
            {
                syncResult.Errors.Add("No SPUTIL placeholders found on this page.");
                return syncResult;
            }

            System.Diagnostics.Debug.WriteLine(
                $"[SyncProperties] Found {placeholders.Count} placeholder(s).");

            // ── Step 2: Get current WebParts on target ────────────────────────
            var targetSnapshot = await GetPageSnapshotAsync(targetSiteUrl, targetPageRelativeUrl);

            // Build lookup: Title → list of WebPartSnapshot (handle duplicates)
            var targetByTitle = new Dictionary<string, List<WebPartSnapshot>>(StringComparer.OrdinalIgnoreCase);
            foreach (var tw in targetSnapshot.WebParts)
            {
                if (!targetByTitle.ContainsKey(tw.Title))
                    targetByTitle[tw.Title] = new List<WebPartSnapshot>();
                targetByTitle[tw.Title].Add(tw);
            }

            // ── Step 3: For each placeholder — find match and sync ────────────
            var matchedTargetKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var meta in placeholders)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[SyncProperties] Processing placeholder: '{meta.Title}' pos={meta.Position}");

                // ── Find matching WebPart on target ───────────────────────────
                if (!targetByTitle.TryGetValue(meta.Title, out var candidates))
                {
                    syncResult.SkippedCount++;
                    syncResult.Errors.Add(
                        $"No WebPart named '{meta.Title}' found on target page. " +
                        $"Please add it manually then run Sync again.");
                    continue;
                }

                // For duplicates — pick closest position, skip already matched
                var match = candidates
                    .Where(c => !matchedTargetKeys.Contains(c.StorageKey))
                    .OrderBy(c => Math.Abs(c.VisualPosition - meta.Position))
                    .FirstOrDefault();

                if (match == null)
                {
                    syncResult.SkippedCount++;
                    syncResult.Errors.Add(
                        $"All WebParts named '{meta.Title}' are already matched. " +
                        $"Add another instance manually.");
                    continue;
                }

                matchedTargetKeys.Add(match.StorageKey);

                // ── Get source ExportXml ──────────────────────────────────────
                string exportXml = await GetWebPartExportXmlAsync(
                    meta.SiteUrl, meta.PageUrl, meta.StorageKey);

                if (string.IsNullOrEmpty(exportXml))
                {
                    syncResult.SkippedCount++;
                    syncResult.Errors.Add(
                        $"Could not download settings for '{meta.Title}' from source. " +
                        $"Source page: {meta.SiteUrl}{meta.PageUrl}");
                    continue;
                }

                // ── Parse properties from ExportXml ──────────────────────────
                var propsToSync = ParsePropertiesFromExportXml(exportXml);

                // ── Apply to target WebPart ───────────────────────────────────
                try
                {
                    var request = new WebPartUpdateRequest
                    {
                        StorageKey        = match.StorageKey,
                        PropertiesToUpdate = propsToSync
                    };

                    await UpdateWebPartAsync(targetSiteUrl, targetPageRelativeUrl, request);

                    // ── Remove placeholder from PublishingContent ─────────────
                    await RemovePlaceholderFromPageAsync(
                        targetSiteUrl, targetPageRelativeUrl, meta.TargetZoneKey);

                    syncResult.SyncedCount++;
                    System.Diagnostics.Debug.WriteLine(
                        $"[SyncProperties] Synced '{meta.Title}' → StorageKey={match.StorageKey}");
                }
                catch (Exception ex)
                {
                    _log.Caller().Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                    syncResult.Errors.Add($"Error syncing '{meta.Title}': {ex.Message}");
                    System.Diagnostics.Debug.WriteLine(
                        $"[SyncProperties] Error for '{meta.Title}': {ex}");
                }
            }

            System.Diagnostics.Debug.WriteLine(
                $"[SyncProperties] Done. {syncResult.Summary}");

            return syncResult;
        }

        /// <summary>
        /// Parses custom property values from a .webpart XML string.
        /// Skips system properties (AllowClose, ChromeType, etc.).
        /// </summary>
        private static Dictionary<string, string> ParsePropertiesFromExportXml(string xml)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var skipProps = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "AllowClose","AllowConnect","AllowEdit","AllowHide",
                "AllowMinimize","AllowZoneChange","AuthorizationFilter",
                "CatalogIconImageUrl","ChromeState","ChromeType",
                "Direction","ExportMode","HelpMode","HelpUrl",
                "Hidden","ImportErrorMessage","TitleIconImageUrl","TitleUrl",
                "Title","Description"
            };

            try
            {
                var doc = System.Xml.Linq.XDocument.Parse(xml);
                var ns  = doc.Root?.GetDefaultNamespace();

                var properties = doc.Descendants()
                    .Where(e => e.Name.LocalName == "property");

                foreach (var prop in properties)
                {
                    string name = prop.Attribute("name")?.Value ?? "";
                    string val  = prop.Value ?? "";

                    if (string.IsNullOrEmpty(name)) continue;
                    if (skipProps.Contains(name)) continue;

                    result[name] = val;
                }
            }
            catch (Exception ex)
            {
                _log.Caller().Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                System.Diagnostics.Debug.WriteLine($"[ParseExportXml] Error: {ex.Message}");
            }

            return result;
        }

        /// <summary>
        /// Removes the placeholder div with the given ZoneKey from PublishingPageContent.
        /// Called after successful Sync of a WebPart.
        /// </summary>
        private async Task RemovePlaceholderFromPageAsync(
            string siteUrl,
            string pageRelativeUrl,
            string zoneKey)
        {
            if (string.IsNullOrEmpty(zoneKey)) return;

            using var ctx = await GetContextAsync(siteUrl);
            var pageFile = ctx.Web.GetFileByServerRelativeUrl(pageRelativeUrl);
            ctx.Load(pageFile, f => f.ListItemAllFields);
            await Task.Run(() => ctx.ExecuteQuery());

			await SafeCheckOutAsync(ctx, pageFile);

			bool checkedIn = false;
			try
			{
				ctx.Load(pageFile.ListItemAllFields);
				await Task.Run(() => ctx.ExecuteQuery());

				var fields  = pageFile.ListItemAllFields;
				string html = fields["PublishingPageContent"]?.ToString() ?? "";

				fields["PublishingPageContent"] = RemovePlaceholderFromHtml(html, zoneKey);
				fields.Update();
				await Task.Run(() => ctx.ExecuteQuery());

				await CheckInAndPublishAsync(ctx, pageFile, $"Removed placeholder {zoneKey}");
				checkedIn = true;
			}
			finally
			{
				if (!checkedIn)
				{
					try
					{
						pageFile.UndoCheckOut();
						await Task.Run(() => ctx.ExecuteQuery());
						_logPage.Caller().Warning(
							"RemovePlaceholderFromPageAsync: operation failed — checkout discarded for {Page}",
							pageRelativeUrl);
					}
					catch (Exception cleanupEx)
					{
						_logPage.Caller().Error(cleanupEx,
							"RemovePlaceholderFromPageAsync: failed to undo checkout after error — page may remain checked out: {Page}",
							pageRelativeUrl);
					}
				}
			}

        }


        // ═══════════════════════════════════════════════════════════════════════
        //  GetPageRelativeUrlAsync
        //  Returns the server-relative URL for a page by filename, optionally
        //  inside a subfolder of the Pages library.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<string> GetPageRelativeUrlAsync(
            string siteUrl,
            string pageName,
            string subfolderPath = "")
        {
            using var ctx = await GetContextAsync(siteUrl);
            ctx.Load(ctx.Web, w => w.ServerRelativeUrl);
            await Task.Run(() => ctx.ExecuteQuery());

            // Shared with PageExists / Delete / Rename: handles the .aspx suffix and
            // subfolders of any depth. Without it the URL always pointed at the Pages
            // root, so Compare and Sync silently looked at a page that does not exist
            // and failed later with "The property or field has not been initialized".
            return BuildPageRelativeUrl(ctx.Web.ServerRelativeUrl, pageName, subfolderPath);
        }

       /// <summary>
        /// Looks up a page layout in the Master Page Gallery of the target site
        /// collection by file name (e.g. "NoMenuPage.aspx").
        /// Returns null when the layout is not present.
        /// </summary>
        /// <remarks>
        /// The gallery is resolved by URL rather than by title on purpose: the
        /// display name is localized (for example "גלריית דפי אב" on Hebrew sites),
        /// so GetByTitle would fail there.
        /// </remarks>

        private async Task<ListItem?> FindPageLayoutItemAsync(
            ClientContext ctx,
            string layoutFileName)
        {
            if (string.IsNullOrWhiteSpace(layoutFileName)) return null;

            var rootWeb = ctx.Site.RootWeb;
            ctx.Load(rootWeb, w => w.ServerRelativeUrl);
            await Task.Run(() => ctx.ExecuteQuery());

            string galleryUrl = rootWeb.ServerRelativeUrl.TrimEnd('/') + "/_catalogs/masterpage";
            var gallery = rootWeb.GetList(galleryUrl);

            // Scope='RecursiveAll' — layouts may live in subfolders of the gallery.
            var query = new CamlQuery
            {
                ViewXml =
                    "<View Scope='RecursiveAll'><Query><Where><Eq>" +
                    "<FieldRef Name='FileLeafRef' />" +
                    "<Value Type='Text'>" +
                    System.Security.SecurityElement.Escape(layoutFileName) +
                    "</Value>" +
                    "</Eq></Where></Query><RowLimit>1</RowLimit></View>"
            };

            var items = gallery.GetItems(query);
            ctx.Load(items, icol => icol.Include(
                i => i.Id,
                i => i["FileLeafRef"],
                i => i["FileRef"]));
            await Task.Run(() => ctx.ExecuteQuery());

            if (items.Count == 0) return null;

            // Return the item by a stable object path instead of items[0].
            // The path of a query result is "List.GetItems(caml)[0]" — the server
            // re-executes the CAML query every time that path is resolved in a later
            // batch. AddPublishingPage resolves PageLayoutListItem in its own batch,
            // so a re-resolvable-by-Id path is safer and cheaper.
            // The returned item is intentionally not loaded: callers only use it as
            // an object path (PageLayoutListItem) or test it for null.
            return gallery.GetItemById(items[0].Id);
        }

        /// <summary>
        /// Extracts the layout file name from a server-relative layout URL,
        /// undoing the percent-encoding produced by Uri.AbsolutePath.
        /// </summary>
        private static string GetLayoutFileName(string layoutRelativeUrl)
        {
            if (string.IsNullOrWhiteSpace(layoutRelativeUrl)) return string.Empty;

            int lastSlash = layoutRelativeUrl.LastIndexOf('/');
            string name = lastSlash >= 0
                ? layoutRelativeUrl.Substring(lastSlash + 1)
                : layoutRelativeUrl;

            return Uri.UnescapeDataString(name);
        }
		

        /// <summary>
        /// Checks whether a page layout with the given file name exists in the
        /// Master Page Gallery of the site collection that owns <paramref name="siteUrl"/>.
        /// Used as a pre-flight check before the expensive snapshot step.
        /// </summary>
        public async Task<bool> PageLayoutExistsAsync(string siteUrl, string layoutFileName)
        {
            if (string.IsNullOrWhiteSpace(layoutFileName)) return false;

            return await Task.Run(async () =>
            {
                using var ctx = await GetContextAsync(siteUrl);
                var item = await FindPageLayoutItemAsync(ctx, layoutFileName);

                bool found = item != null;
                _logPage.Information(
                    "PageLayoutExists: '{Layout}' on {Site} → {Found}",
                    layoutFileName, siteUrl, found);

                return found;
            });
        }
		/*
		/// <summary>
        /// Reads a CSOM property, returning a marker instead of throwing when the
        /// server did not initialize it despite the Include.
        /// </summary>
        private static string SafeRead(Func<string> read)
        {
            try   { return read(); }
            catch (Exception ex) { return "<n/a:" + ex.GetType().Name + ">"; }
        }
		*/
		private static T SafeValue<T>(Func<T> read, T fallback)
        {
            try   { return read(); }
            catch { return fallback; }
        }

		/// <summary>
		/// Builds a Content Editor Web Part's import XML carrying a fixed warning
		/// message — used as a visible placeholder in a layout zone when the real
		/// WebPart could not be copied. ContentEditorWebPart is core SharePoint,
		/// present on every farm without any Feature activation; only its Assembly
		/// version differs by SharePoint generation (major version only — CU-level
		/// patches don't change it, SharePoint's own web.config binding redirects
		/// handle that).
		/// </summary>
		private static string BuildContentEditorWarningXml(int serverMajorVersion, string htmlMessage)
		{
			string assemblyVersion = $"{serverMajorVersion}.0.0.0";

				return $@"<?xml version=""1.0"" encoding=""utf-8""?>
			<WebPart xmlns=""http://schemas.microsoft.com/WebPart/v2"">
				<Assembly>Microsoft.SharePoint, Version={assemblyVersion}, Culture=neutral, PublicKeyToken=71e9bce111e9429c</Assembly>
				<TypeName>Microsoft.SharePoint.WebPartPages.ContentEditorWebPart</TypeName>
				<Title>A required Web Part is missing from this site.</Title>
				<Description>Warning placeholder for missing web part</Description>
				<PartImageLarge>/_layouts/15/images/mscontl.gif</PartImageLarge>
				<ContentLink xmlns=""http://schemas.microsoft.com/WebPart/v2/ContentEditor"" />
				<Content xmlns=""http://schemas.microsoft.com/WebPart/v2/ContentEditor""><![CDATA[{htmlMessage}]]></Content>				
			</WebPart>";
		}
		
		// ═══════════════════════════════════════════════════════════════════════
        //  EnsureSubfolderAsync
        //  Creates the subfolder hierarchy under the Pages library if it doesn't
        //  exist. Returns the deepest Folder object.
        //  subfolderPath examples: "Dean"  /  "FacultyAdmin/Sub"
        // ═══════════════════════════════════════════════════════════════════════
        private async Task<Folder> EnsureSubfolderAsync(
            ClientContext ctx,
            string pagesRootServerRelativeUrl,
            string subfolderPath)
        {
            var parts = subfolderPath.Split(
                new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries);

            string currentPath = pagesRootServerRelativeUrl;
            Folder folder      = ctx.Web.GetFolderByServerRelativeUrl(currentPath);
            ctx.Load(folder);
            await Task.Run(() => ctx.ExecuteQuery());

            foreach (var part in parts)
            {
                string childPath = currentPath.TrimEnd('/') + "/" + part;
                try
                {
                    var child = ctx.Web.GetFolderByServerRelativeUrl(childPath);
                    ctx.Load(child, f => f.Name);
                    await Task.Run(() => ctx.ExecuteQuery());
                    // Folder already exists — continue
                    folder      = child;
                    currentPath = childPath;
                }
                catch (ServerException ex) when (
                    ex.ServerErrorTypeName == "System.IO.FileNotFoundException" ||
                    ex.Message.Contains("does not exist") ||
                    ex.Message.Contains("FileNotFoundException"))
                {
                    // Folder doesn't exist — create it
                    folder.Folders.Add(part);
                    await Task.Run(() => ctx.ExecuteQuery());

                    folder = ctx.Web.GetFolderByServerRelativeUrl(childPath);
                    ctx.Load(folder, f => f.Name, f => f.ServerRelativeUrl);
                    await Task.Run(() => ctx.ExecuteQuery());

                    currentPath = childPath;
                    System.Diagnostics.Debug.WriteLine(
                        $"[EnsureSubfolder] Created: {childPath}");
                }
            }

            return folder;
        }
    }
}
