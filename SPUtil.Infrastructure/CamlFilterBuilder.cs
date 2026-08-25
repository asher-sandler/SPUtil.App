using System;

namespace SPUtil.Infrastructure
{
    /// <summary>
    /// Builds a CAML &lt;Where&gt; clause for the single-condition Filter dialog
    /// (List100ViewModel). Supports one field at a time — no multi-field AND/OR,
    /// matching the v1 scope agreed for the Filter feature.
    /// </summary>
    public static class CamlFilterBuilder
    {
        /// <summary>
        /// Builds the full &lt;Where&gt;...&lt;/Where&gt; clause.
        /// value2 is required only when operatorName == "Between".
        /// </summary>
        public static string BuildWhereClause(string fieldName, string operatorName, string value1, string? value2 = null)
        {
            if (string.IsNullOrWhiteSpace(fieldName))
                throw new ArgumentException("Field name is required.", nameof(fieldName));

            if (string.IsNullOrWhiteSpace(value1))
                throw new ArgumentException("Value is required.", nameof(value1));

            string inner = operatorName switch
            {
                "Equals"       => BuildSingle("Eq", fieldName, value1),
                "Greater than" => BuildSingle("Gt", fieldName, value1),
                "Less than"    => BuildSingle("Lt", fieldName, value1),
                "Contains"     => BuildSingle("Contains", fieldName, value1),
                "Between"      => BuildBetween(fieldName, value1, value2),
                _ => throw new ArgumentException($"Unsupported operator: {operatorName}", nameof(operatorName))
            };

            return $"<Where>{inner}</Where>";
        }

        private static string BuildBetween(string fieldName, string value1, string? value2)
        {
            if (string.IsNullOrWhiteSpace(value2))
                throw new ArgumentException("Between requires a second (To) value.", nameof(value2));

            // CAML has no native "Between" tag — built as two range
            // conditions combined with And, per the model agreed earlier.
            string geq = BuildSingle("Geq", fieldName, value1);
            string leq = BuildSingle("Leq", fieldName, value2);
            return $"<And>{geq}{leq}</And>";
        }

        private static string BuildSingle(string camlOp, string fieldName, string value)
        {
            string typeAttr = FieldValueType(fieldName);
            string escaped  = System.Security.SecurityElement.Escape(value);
            return $"<{camlOp}><FieldRef Name='{fieldName}'/><Value Type='{typeAttr}'>{escaped}</Value></{camlOp}>";
        }

        /// <summary>Maps a fixed v1 field name to its CAML Value/@Type.</summary>
        private static string FieldValueType(string fieldName) => fieldName switch
        {
            "ID"       => "Number",
            "Modified" => "DateTime",
            _          => "Text"   // Title
        };
    }
}