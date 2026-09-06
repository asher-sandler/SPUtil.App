using System.Collections.Generic;

namespace SPUtil.Infrastructure
{
    // ═══════════════════════════════════════════════════════════════════════════
    //  WebPartSavePropertiesResult — per-property outcome of saving edits to a
    //  single WebPart's properties (SaveSingleWebPartPropertiesAsync).
    //
    //  Unlike UpdateWebPartAsync/UpdateAllWebPartsAsync (which return plain Task
    //  and only log per-property failures), this DTO lets the caller tell the
    //  user exactly which properties were written and which were not — needed
    //  because some properties are known to silently fail on the server side
    //  (enum-backed custom properties such as textAlign/textDirection — see the
    //  comment above UpdateAllWebPartsAsync in SharePointPageService.cs).
    // ═══════════════════════════════════════════════════════════════════════════
    public class WebPartSavePropertiesResult
    {
        /// <summary>Property names that were written and confirmed via ExecuteQuery.</summary>
        public List<string> Succeeded { get; set; } = new();

        /// <summary>Property name -> failure reason (ex.Message), for properties that were not saved.</summary>
        public Dictionary<string, string> Failed { get; set; } = new();

        /// <summary>True when every requested property was saved successfully.</summary>
        public bool IsFullSuccess => Failed.Count == 0;
    }
}
