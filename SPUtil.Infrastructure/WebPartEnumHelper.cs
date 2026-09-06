using System;
using System.Collections.Generic;
using System.Linq;

namespace SPUtil.Infrastructure
{
    /// <summary>
    /// Decode table for the two WebPart properties confirmed to carry the same meaning
    /// across different controls: "textAlign" and "textDirection" (verified against the
    /// native SharePoint tool pane for "Dynamic Form - v 2.0"). Everything else is edited
    /// as plain text in the custom properties editor — other enum-like properties
    /// (formMode, itemLoadBy, fontName, etc.) turned out to be control-specific and are
    /// intentionally NOT generalized here; a wrong guess would silently misconfigure
    /// a control, so plain text is the safer default until a property is proven global.
    /// </summary>
    public static class WebPartEnumHelper
    {
        private static readonly Dictionary<string, Dictionary<int, string>> _known =
            new(StringComparer.OrdinalIgnoreCase)
        {
            ["textAlign"] = new()
            {
                [0] = "right",
                [1] = "left"
            },
            ["textDirection"] = new()
            {
                [0] = "rtl",
                [1] = "ltr"
            }
        };

        /// <summary>True if this property name is one of the known global enums.</summary>
        public static bool IsKnownEnum(string propertyName) =>
            !string.IsNullOrEmpty(propertyName) && _known.ContainsKey(propertyName);

        /// <summary>Display labels for a dropdown, in raw-value order. Empty if not a known enum.</summary>
        public static IReadOnlyList<string> GetOptions(string propertyName)
        {
            if (string.IsNullOrEmpty(propertyName) || !_known.TryGetValue(propertyName, out var values))
                return Array.Empty<string>();

            return values.OrderBy(kv => kv.Key).Select(kv => kv.Value).ToList();
        }

        /// <summary>Raw "0"/"1" value -> display label ("right"/"rtl"...). False if unknown.</summary>
        public static bool TryGetLabel(string propertyName, string rawValue, out string label)
        {
            label = rawValue;

            if (!IsKnownEnum(propertyName) || !int.TryParse(rawValue, out int intValue))
                return false;

            if (!_known[propertyName].TryGetValue(intValue, out var text))
                return false;

            label = text;
            return true;
        }

        /// <summary>Display label ("right"/"rtl"...) -> raw "0"/"1" value, for writing back via UpdateWebPartAsync.</summary>
        public static bool TryGetRawValue(string propertyName, string label, out string rawValue)
        {
            rawValue = label;

            if (!IsKnownEnum(propertyName))
                return false;

            var match = _known[propertyName]
                .FirstOrDefault(kv => string.Equals(kv.Value, label, StringComparison.OrdinalIgnoreCase));

            if (match.Value == null)
                return false;

            rawValue = match.Key.ToString();
            return true;
        }
    }
}