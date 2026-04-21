using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.WebParts;
using SPUtil.Infrastructure;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SPUtil.Services
{
    /// <summary>
    /// Partial class — WebPart management operations that enrich the basic
    /// GetWebPartsAsync result with VisualPosition data.
    ///
    /// SharePointService.cs is NOT modified. This file adds a new method
    /// GetWebPartsWithPositionAsync that wraps the existing GetWebPartsAsync
    /// and resolves the visual order from PublishingPageContent.
    ///
    /// PagesViewModel should call GetWebPartsWithPositionAsync instead of
    /// GetWebPartsAsync when it needs VisualPosition populated.
    /// </summary>
    public partial class SharePointService
    {
        // ═══════════════════════════════════════════════════════════════════════
        //  GetWebPartsWithPositionAsync
        //  Calls the existing GetWebPartsAsync, then resolves VisualPosition
        //  for each WebPart by reading PublishingPageContent and the rendered HTML.
        //  Does not modify SharePointService.cs.
        // ═══════════════════════════════════════════════════════════════════════
        public async Task<List<SPWebPartData>> GetWebPartsWithPositionAsync(
            string siteUrl,
            string fileRelativeUrl)
        {
            // Step 1: Get WebParts via existing method (unmodified)
            var webParts = await GetWebPartsAsync(siteUrl, fileRelativeUrl);
            if (!webParts.Any()) return webParts;

            // Step 2: Build StorageKey → VisualPosition mapping
            var storageKeyToPosition = new Dictionary<string, int>(
                StringComparer.OrdinalIgnoreCase);

            try
            {
                // Read PublishingPageContent via CSOM to get ZoneKey order
                using var ctx = await GetContextAsync(siteUrl);
                var file = ctx.Web.GetFileByServerRelativeUrl(fileRelativeUrl);
                ctx.Load(file, f => f.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());
                ctx.Load(file.ListItemAllFields);
                await Task.Run(() => ctx.ExecuteQuery());

                string pubHtml = file.ListItemAllFields["PublishingPageContent"]
                                     ?.ToString() ?? "";

                // Get ZoneKeys in visual order from PublishingContent
                var zoneKeysInOrder = ParseZoneKeysInOrderStatic(pubHtml);
                if (!zoneKeysInOrder.Any())
                {
                    System.Diagnostics.Debug.WriteLine(
                        "[GetWebPartsWithPosition] No ZoneKeys found in PublishingContent.");
                    return SortByPosition(webParts);
                }

                // Get webpartid → webpartid2 mapping from rendered HTML
                string rendered = await FetchRenderedHtmlAsync(siteUrl, fileRelativeUrl);
                var zoneKeyToStorageKey = ParseZoneKeyToStorageKeyStatic(rendered);

                // Map: ZoneKey order → StorageKey → position
                int pos = 1;
                foreach (var zk in zoneKeysInOrder)
                {
                    if (zoneKeyToStorageKey.TryGetValue(zk, out var sk))
                        storageKeyToPosition[sk] = pos;
                    pos++;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[GetWebPartsWithPosition] VisualPosition resolution failed: {ex.Message}");
                return SortByPosition(webParts);
            }

            // Step 3: Assign VisualPosition to each WebPart
            foreach (var wp in webParts)
            {
                if (storageKeyToPosition.TryGetValue(wp.StorageKey, out int vp))
                    wp.VisualPosition = vp;
                // else VisualPosition stays 0 — WebPart is in a named zone
            }

            return SortByPosition(webParts);
        }

        /// <summary>Sort: positioned WPs first (by position), then unpositioned (zone WPs)</summary>
        private static List<SPWebPartData> SortByPosition(List<SPWebPartData> list) =>
            list.OrderBy(w => w.VisualPosition == 0 ? int.MaxValue : w.VisualPosition)
                .ToList();

        // ── Static parsing helpers ─────────────────────────────────────────────
        // These duplicate the logic from the private methods in SharePointPageService
        // so we don't need to change their visibility.

        /// <summary>
        /// Fetches rendered page HTML via HTTP GET with NormalizeUrl applied.
        /// Standalone version — does not depend on SharePointPageService private members.
        /// </summary>
        private async Task<string> FetchRenderedHtmlAsync(string siteUrl, string pageRelativeUrl)
        {
            string hostRoot = "https://" + new Uri(siteUrl).Host;
            string fullUrl  = SPUtil.Infrastructure.SPUsingUtils.NormalizeUrl(
                hostRoot + pageRelativeUrl);

            var handler = new System.Net.Http.HttpClientHandler
            {
                Credentials = GetCredentials()
            };
            using var http = new System.Net.Http.HttpClient(handler);
            http.DefaultRequestHeaders.Add("Accept", "text/html");

            var response = await http.GetAsync(fullUrl);
            if (!response.IsSuccessStatusCode)
            {
                System.Diagnostics.Debug.WriteLine(
                    $"[FetchRenderedHtml] HTTP {(int)response.StatusCode} for {fullUrl}");
                return string.Empty;
            }
            return await response.Content.ReadAsStringAsync();
        }

        /// <summary>
        /// Parses webpartid (StorageKey) → webpartid2 (ZoneKey) pairs from rendered HTML.
        /// Returns ZoneKey → StorageKey dictionary.
        /// </summary>
        private static Dictionary<string, string> ParseZoneKeyToStorageKeyStatic(string html)
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            if (string.IsNullOrEmpty(html)) return result;

            var matches = Regex.Matches(
                html,
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

        /// <summary>
        /// Extracts ZoneKey GUIDs from PublishingPageContent in visual order
        /// (order of appearance in the HTML string).
        /// </summary>
        private static List<string> ParseZoneKeysInOrderStatic(string publishingHtml)
        {
            var result  = new List<string>();
            if (string.IsNullOrEmpty(publishingHtml)) return result;

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
    }
}
