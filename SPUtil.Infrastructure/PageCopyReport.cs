using System;
using System.Collections.Generic;
using System.Linq;

namespace SPUtil.Infrastructure
{
    /// <summary>
    /// Outcome of a single WebPart during a page copy.
    /// </summary>
    public enum WebPartCopyStatus
    {
        /// <summary>Added to the target page as intended.</summary>
        Ok,

        /// <summary>No attempt was made — there was nothing to add
        /// (typically an empty ExportXml from the source).</summary>
        Skipped,

        /// <summary>An attempt was made and it failed.</summary>
        Failed
    }

    /// <summary>
    /// Where a WebPart was placed on the target page.
    /// </summary>
    public enum WebPartPlacement
    {
        /// <summary>A named zone declared by the page layout (Header, RightColumn…).</summary>
        LayoutZone,

        /// <summary>Inline in PublishingPageContent via the wpz pseudo-zone.</summary>
        PageContent
    }

    /// <summary>
    /// Result of copying one WebPart. Carries enough context to identify the
    /// WebPart on the source page — Title alone is not sufficient when a page
    /// holds several instances of the same web part.
    /// </summary>
    public class WebPartCopyEntry
    {
        public WebPartCopyStatus Status    { get; set; }
        public WebPartPlacement  Placement { get; set; }

        /// <summary>WebPart title as shown on the source page.</summary>
        public string Title { get; set; } = string.Empty;

        /// <summary>Target zone id: a layout zone name, or "wpz" for page content.</summary>
        public string ZoneId { get; set; } = string.Empty;

        /// <summary>
        /// Position inside the zone for layout zones, or visual position in
        /// PublishingPageContent for inline WebParts.
        /// </summary>
        public int Position { get; set; }

        /// <summary>StorageKey on the SOURCE page — lets the user locate the
        /// WebPart there via ?contents=1 or a previous report.</summary>
        public string SourceStorageKey { get; set; } = string.Empty;

        /// <summary>StorageKey assigned on the target page. Empty unless Status is Ok.</summary>
        public string TargetStorageKey { get; set; } = string.Empty;

        /// <summary>Why the WebPart was skipped or failed. Empty when Status is Ok.</summary>
        public string Reason { get; set; } = string.Empty;
    }

    /// <summary>
    /// Full report of a page copy operation: what was created, where, and what
    /// did not make it. Produced by CreatePageFromSnapshotAsync and rendered for
    /// the user when anything went wrong.
    /// </summary>
    public class PageCopyReport
    {
        public string SourcePageUrl { get; set; } = string.Empty;
        public string TargetPageUrl { get; set; } = string.Empty;
        public string LayoutName    { get; set; } = string.Empty;
        public DateTime CopyTime    { get; set; } = DateTime.Now;

        public List<WebPartCopyEntry> Entries { get; set; } = new();

        public int TotalCount   => Entries.Count;
        public int OkCount      => Entries.Count(e => e.Status == WebPartCopyStatus.Ok);
        public int SkippedCount => Entries.Count(e => e.Status == WebPartCopyStatus.Skipped);
        public int FailedCount  => Entries.Count(e => e.Status == WebPartCopyStatus.Failed);

        public int InZonesCount => Entries.Count(e =>
            e.Status == WebPartCopyStatus.Ok && e.Placement == WebPartPlacement.LayoutZone);

        public int InContentCount => Entries.Count(e =>
            e.Status == WebPartCopyStatus.Ok && e.Placement == WebPartPlacement.PageContent);

        /// <summary>True when every WebPart was copied — the report needs no attention.</summary>
        public bool IsClean => SkippedCount == 0 && FailedCount == 0;

        /// <summary>One-line summary for the Copy Complete dialog.</summary>
        public string Summary => IsClean
            ? $"WebParts copied: {OkCount} ({InZonesCount} in zones, {InContentCount} in content)"
            : $"WebParts: {TotalCount} total — {OkCount} added, {SkippedCount} skipped, {FailedCount} failed";
    }
}