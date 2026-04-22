using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SPUtil.Infrastructure
{
    public class SPFieldData
    {
        public string Id { get; set; } = string.Empty;
        public string Title { get; set; } = string.Empty;
        public string InternalName { get; set; } = string.Empty;
        public string TypeAsString { get; set; } = string.Empty;
        public bool Required { get; set; }
        public string? LookupList { get; set; }
    }

    public class SPFileData
    {
        public bool IsSelected { get; set; }
        public string Name { get; set; } = string.Empty;
        public string FullPath { get; set; } = string.Empty;
        public long Size { get; set; }
        public DateTime Modified { get; set; }
        public bool IsFolder { get; set; }
    }

    public class SPWebPartData
    {
        /// <summary>
        /// WebPart instance object ID — internal identifier within the WebPart object model.
        /// Comes from WebPart.Id (not the same as StorageKey).
        /// </summary>
        public string Id { get; set; } = string.Empty;

        /// <summary>
        /// StorageKey — the GUID that SharePoint writes into the ms-rte-wpbox placeholder
        /// inside PublishingPageContent. This is WebPartDefinition.Id, NOT WebPart.Id.
        /// Use this to match a WebPart in Properties output to a div in the page HTML.
        /// </summary>
        public string StorageKey { get; set; } = string.Empty;

        public string Title { get; set; } = string.Empty;
        public string Type  { get; set; } = string.Empty;
        public string ZoneId { get; set; } = string.Empty;

        /// <summary>Visual position on the page (1-based). Used for matching duplicates.</summary>
        public int VisualPosition { get; set; }

        public Dictionary<string, string> Properties { get; set; } = new();
    }
	public class FieldInfo
	{
		public string Id { get; set; } = string.Empty;
		public string Name { get; set; } = string.Empty;           // InternalName
		public string StaticName { get; set; } = string.Empty;     // Статическое имя
		public string DisplayName { get; set; } = string.Empty;
		public string FieldType { get; set; } = string.Empty;
		public string Required { get; set; } = string.Empty;
		public string EnforceUniqueValues { get; set; } = "False";

		// Properties, вызывавшие ошибки CS1061:
		public int MaxLength { get; set; } = 255;                  // Для типа Text
		public int NumLines { get; set; } = 6;                     // Для типа Note
		public int? Decimals { get; set; }                         // Для Number/Currency
		public string Formula { get; set; } = string.Empty;        // Для Calculated
		public string ResultType { get; set; } = "Text";           // Для Calculated
		public string DefaultValue { get; set; } = string.Empty;   // Значение по умолчанию

		// Для Lookup
		public string LookupListId { get; set; } = string.Empty;
		public string LookupListName { get; set; } = string.Empty;
		public string LookupWebId { get; set; } = string.Empty;
		public string LookupFieldName { get; set; } = string.Empty;
		public string PrimaryFieldId { get; set; } = string.Empty; // Это ID родителя со старого сайта
		public string FieldRef { get; set; } = string.Empty;       // А это будет новый ID для XML
		public bool IsDependentLookup { get; set; } = false;

        // Для Choice
        public string Format { get; set; } = string.Empty;
		public List<string> Choices { get; set; } = new List<string>();

		// Для Note (RichText)
		public string RichText { get; set; } = string.Empty;
		public string RichTextMode { get; set; } = string.Empty;
		public string IsolateStyles { get; set; } = string.Empty;
		public string SchemaXml { get; set; } = string.Empty;

		// Для Number / Currency
		public double? MinValue { get; set; }
		public double? MaxValue { get; set; }

		/// <summary>
		/// Generates SharePoint Field XML.
		/// Supports: Text, Note, Choice, MultiChoice, Number, Currency, DateTime, Lookup, Boolean, User, URL, Calculated.
		/// </summary>
		public string BuildXml()
		{
			// Вместо StringBuilder для сложной логики XML лучше использовать XElement, 
			// но адаптируем ваш текущий BuildXml под новые поля:
			
			System.Text.StringBuilder sb = new System.Text.StringBuilder();

			// Базовые атрибуты
			sb.Append($"<Field Type='{FieldType}' DisplayName='{DisplayName}' ");

			if (!string.IsNullOrEmpty(Name))
			{
				//sb.Append($"Name='{Name}' ");
				//sb.Append($"StaticName='{(!string.IsNullOrEmpty(StaticName) ? StaticName : Name)}' ");
			}
			
			if (!string.IsNullOrEmpty(Required)){

				sb.Append($"Required='{Required?.ToString()}' ");
			}

			switch (FieldType)
			{
				case "Text":
					sb.Append($"MaxLength='{(MaxLength > 0 ? MaxLength : 255)}' ");
					break;

				case "TargetTo":
					// Специфические атрибуты для Audience Targeting
					sb.Append("ReadOnly='False' Sortable='False' ");
					break;
					

				case "Note":
					sb.Append($"NumLines='{(NumLines > 0 ? NumLines : 6)}' ");
					if (!string.IsNullOrEmpty(RichText))
					{
						sb.Append($"RichText='{RichText.ToString()}' ");
						if (!string.IsNullOrEmpty(RichTextMode)) sb.Append($"RichTextMode='{RichTextMode}' ");
						if (!string.IsNullOrEmpty(IsolateStyles)) sb.Append($"IsolateStyles='{IsolateStyles.ToUpper()}' ");
					}
					break;

				case "Number":
				case "Currency":
					if (MinValue.HasValue) sb.Append($"Min='{MinValue.Value}' ");
					if (MaxValue.HasValue) sb.Append($"Max='{MaxValue.Value}' ");
					if (Decimals.HasValue) sb.Append($"Decimals='{Decimals.Value}' ");
					break;

				case "URL":
					sb.Append($"Format='{(!string.IsNullOrEmpty(Format) ? Format : "Hyperlink")}' ");
					break;

				case "DateTime":
					sb.Append($"Format='{(!string.IsNullOrEmpty(Format) ? Format : "DateOnly")}' ");
					break;

				case "Calculated":
					sb.Append($"ResultType='{(!string.IsNullOrEmpty(ResultType) ? ResultType : "Text")}' ");
					// ПРАВИЛЬНО: Format как атрибут
					if (!string.IsNullOrEmpty(Format))
					{
						sb.Append($"Format='{Format}' ");
					}
					break;

				case "Choice":
				case "MultiChoice":
					sb.Append($"Format='{(!string.IsNullOrEmpty(Format) ? Format : "Dropdown")}' ");
					break;

				case "Lookup":
				case "LookupMulti":
					if (!string.IsNullOrEmpty(LookupListId))
					{
						string listId = LookupListId.StartsWith("{") ? LookupListId : $"{{{LookupListId}}}";
						sb.Append($"List='{listId}' ");
					}
					if (!string.IsNullOrEmpty(LookupWebId))
					{
						string webId = LookupWebId.StartsWith("{") ? LookupWebId : $"{{{LookupWebId}}}";
						sb.Append($"WebId='{webId}' ");
					}
					sb.Append($"ShowField='{(!string.IsNullOrEmpty(LookupFieldName) ? LookupFieldName : "Title")}' ");
					if (FieldType == "LookupMulti") sb.Append("Mult='TRUE' ");
					break;
			}

			sb.Append(">");

			// Вложенные элементы
			if (FieldType == "Calculated" && !string.IsNullOrEmpty(Formula))
			{
				sb.Append($"<Formula>{Formula}</Formula>");
			}

			if ((FieldType == "Choice" || FieldType == "MultiChoice") && Choices != null && Choices.Any())
			{
				sb.Append("<CHOICES>");
				foreach (var choice in Choices)
				{
					sb.Append($"<CHOICE>{choice}</CHOICE>");
				}
				sb.Append("</CHOICES>");
			}

			if (!string.IsNullOrEmpty(DefaultValue))
			{
				sb.Append($"<Default>{DefaultValue}</Default>");
			}
			else if (FieldType == "Boolean")
			{
				sb.Append("<Default>0</Default>");
			}

			sb.Append("</Field>");
			return sb.ToString();
		}
	}
}