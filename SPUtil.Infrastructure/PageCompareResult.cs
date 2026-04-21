using System.Collections.Generic;
using System.Linq;

namespace SPUtil.Infrastructure
{
    // ═══════════════════════════════════════════════════════════════════════════
    //  WebPartPlaceholderMeta
    //  Метаданные источника хранящиеся в JSON-комментарии внутри заглушки:
    //  <!--SPUTIL:{"storageKey":"...","title":"...","position":1,...}-->
    //  SharePoint не удаляет HTML-комментарии при сохранении через CSOM.
    // ═══════════════════════════════════════════════════════════════════════════
    public class WebPartPlaceholderMeta
    {
        public string StorageKey      { get; set; } = string.Empty;  // source
        public string Title           { get; set; } = string.Empty;
        public int    Position        { get; set; }
        public string ZoneId          { get; set; } = string.Empty;
        public string SiteUrl         { get; set; } = string.Empty;  // source site
        public string PageUrl         { get; set; } = string.Empty;  // source page

        // Filled at parse time — identifies the placeholder on target
        public string TargetZoneKey    { get; set; } = string.Empty;
        public string TargetStorageKey { get; set; } = string.Empty;
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  SyncResult — результат SyncPropertiesAsync
    // ═══════════════════════════════════════════════════════════════════════════
    public class SyncResult
    {
        public int          SyncedCount  { get; set; }
        public int          SkippedCount { get; set; }
        public List<string> Errors       { get; set; } = new();
        public bool         IsSuccess    => !Errors.Any() && SkippedCount == 0;
        public string       Summary      =>
            $"Synced: {SyncedCount}  |  Skipped: {SkippedCount}  |  Errors: {Errors.Count}";
    }

    // ═══════════════════════════════════════════════════════════════════════════
    //  PageCompareResult — структурированный diff двух снимков
    // ═══════════════════════════════════════════════════════════════════════════
    public enum WebPartDiffKind { Identical, Added, Removed, Modified }

    public class PropertyDiff
    {
        public string PropertyName { get; set; } = string.Empty;
        public string SourceValue  { get; set; } = string.Empty;
        public string TargetValue  { get; set; } = string.Empty;
    }

    public class WebPartDiff
    {
        public WebPartDiffKind    Kind            { get; set; }
        public string             Title           { get; set; } = string.Empty;
        public int                SourcePosition  { get; set; }
        public int                TargetPosition  { get; set; }
        public string             SourceStorageKey { get; set; } = string.Empty;
        public string             SourceZoneId    { get; set; } = string.Empty;
        public List<PropertyDiff> PropertyDiffs   { get; set; } = new();
    }

    public class PageCompareResult
    {
        public string SourceUrl      { get; set; } = string.Empty;
        public string TargetUrl      { get; set; } = string.Empty;
        public string SourceSiteUrl  { get; set; } = string.Empty;
        public string TargetSiteUrl  { get; set; } = string.Empty;
        public bool   LayoutMatches  { get; set; }
        public string SourceLayout   { get; set; } = string.Empty;
        public string TargetLayout   { get; set; } = string.Empty;
        public int    SourceWebPartCount { get; set; }
        public int    TargetWebPartCount { get; set; }

        public List<WebPartDiff> Diffs { get; set; } = new();

        public bool IsIdentical    => LayoutMatches && Diffs.All(d => d.Kind == WebPartDiffKind.Identical);
        public int  AddedCount     => Diffs.Count(d => d.Kind == WebPartDiffKind.Added);
        public int  RemovedCount   => Diffs.Count(d => d.Kind == WebPartDiffKind.Removed);
        public int  ModifiedCount  => Diffs.Count(d => d.Kind == WebPartDiffKind.Modified);
        public int  IdenticalCount => Diffs.Count(d => d.Kind == WebPartDiffKind.Identical);

        public IEnumerable<WebPartDiff> RemovedWebParts =>
            Diffs.Where(d => d.Kind == WebPartDiffKind.Removed).OrderBy(d => d.SourcePosition);
    }
}
