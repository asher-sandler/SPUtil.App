using Microsoft.SharePoint.Administration;
using Microsoft.SharePoint.Client;
using SPUtil.Infrastructure;
using System;
using System.Collections.Generic;
using System.Xml.Linq;

namespace SPUtil.Services
{
    public class SharePointCloneService
    {
        public string SourceListTitle { get; set; }
        public string TargetListTitle { get; set; }
        public int ListTemplateType { get; set; } = 100; // По умолчанию Custom List
                                                         // Наша временная заглушка для проверки связи

       public async Task<bool> TestCloneServiceConnectionAsync()
        {
            // Просто возвращаем true, чтобы убедиться, что метод вызывается
            return await Task.FromResult(true);
        }
        
		
		/// <summary>
		/// Generates SharePoint Field XML (SchemaXml) based on FieldInfo.
		/// Supported SharePoint Field Types:
		/// - Text: Single line of text
		/// - Note: Multiple lines of text (Plain, Rich, Enhanced Rich)
		/// - Choice: Drop-down/Radio buttons
		/// - MultiChoice: Checkboxes
		/// - Number: Floating point or integer
		/// - Currency: Financial values
		/// - DateTime: Date only or Date & Time
		/// - Lookup: Reference to another list
		/// - Boolean: Yes/No checkbox
		/// - User / UserMulti: Person or Group selection
		/// - URL: Hyperlink or Image
		/// - Calculated: Formula-based fields
		/// - Guid: Unique identifier
		/// </summary>
		public string GenerateFieldXml(FieldInfo field)
		{
			// 1. Create base element with essential attributes
			XElement fieldXml = new XElement("Field",
				new XAttribute("Type", field.FieldType),
				new XAttribute("DisplayName", field.DisplayName)
				
			);

			// Add Internal and Static names
			if (!string.IsNullOrEmpty(field.Required))
			{
				if (field.Required.ToLower() == "true" || field.Required.ToLower() == "false")
				{
					if (field.FieldType != "Calculated")
					{
						fieldXml.Add(new XAttribute("Required", field.Required));
					}
				}
				//fieldXml.Add(new XAttribute("StaticName", field.StaticName ?? field.Name));
			}

            // Add Default Value if exists (generic for most types)
            if (!string.IsNullOrEmpty(field.DefaultValue) && field.FieldType != "Boolean" && field.FieldType != "Calculated")
            {
                fieldXml.Add(new XElement("Default", field.DefaultValue));
            }


            // 2. Specific logic by field type
            switch (field.FieldType)
			{
				case "Text":
					int ml = field.MaxLength > 0 ? field.MaxLength : 255;
					fieldXml.Add(new XAttribute("MaxLength", ml));
					break;

				case "Note":
					if (!string.IsNullOrEmpty(field.RichText))
					{
						fieldXml.Add(new XAttribute("RichText", field.RichText.ToUpper()));
						if (!string.IsNullOrEmpty(field.RichTextMode)) fieldXml.Add(new XAttribute("RichTextMode", field.RichTextMode));
						if (!string.IsNullOrEmpty(field.IsolateStyles)) fieldXml.Add(new XAttribute("IsolateStyles", field.IsolateStyles.ToUpper()));
					}
					int lines = field.NumLines > 0 ? field.NumLines : 6;
					fieldXml.Add(new XAttribute("NumLines", lines));
					break;

				case "Lookup":
				case "LookupMulti":
					if (!string.IsNullOrEmpty(field.LookupListId))
					{
						string formattedListId = field.LookupListId.StartsWith("{") ? field.LookupListId : $"{{{field.LookupListId}}}";
						fieldXml.Add(new XAttribute("List", formattedListId));
					}
                    // ДОБАВЛЯЕМ ЭТО: FieldRef для зависимых полей
                    if (field.IsDependentLookup && !string.IsNullOrEmpty(field.PrimaryFieldId))
                    {
                        
                        
						
						if (!string.IsNullOrEmpty(field.FieldRef))
						{
							string fRef = field.FieldRef.StartsWith("{") ? field.FieldRef : $"{{{field.FieldRef}}}";
                            fieldXml.Add(new XAttribute("FieldRef", fRef));
						}
						else
						{
                            string formattedFieldRef = field.PrimaryFieldId.StartsWith("{") ? field.PrimaryFieldId : $"{{{field.PrimaryFieldId}}}";
                            fieldXml.Add(new XAttribute("FieldRef", formattedFieldRef));
                        }
                        fieldXml.Add(new XAttribute("ReadOnly", "TRUE"));
                    }

                   
                    if (!string.IsNullOrEmpty(field.LookupFieldName))
					{
						fieldXml.Add(new XAttribute("ShowField", field.LookupFieldName));
					}
					if (!string.IsNullOrEmpty(field.LookupWebId))
					{
						string formattedWebId = field.LookupWebId.StartsWith("{") ? field.LookupWebId : $"{{{field.LookupWebId}}}";
						fieldXml.Add(new XAttribute("WebId", formattedWebId));
					}
					if (field.FieldType == "LookupMulti") fieldXml.Add(new XAttribute("Mult", "TRUE"));
					break;

				case "Choice":
				case "MultiChoice":
					fieldXml.Add(new XAttribute("Format", field.Format ?? "Dropdown"));
					if (field.Choices != null && field.Choices.Any())
					{
						XElement choicesElement = new XElement("CHOICES");
						foreach (var choice in field.Choices)
						{
							choicesElement.Add(new XElement("CHOICE", choice));
						}
						fieldXml.Add(choicesElement);
					}
					break;

				case "User":
				case "UserMulti":
					// SelectionGroup: ID of a SharePoint group to limit selection
					// SelectionMode: 0 = People only, 1 = People and Groups
					fieldXml.Add(new XAttribute("SelectionMode", "PeopleAndGroups")); 
					if (field.FieldType == "UserMulti") fieldXml.Add(new XAttribute("Mult", "TRUE"));
					break;

				case "URL":
					fieldXml.Add(new XAttribute("Format", field.Format ?? "Hyperlink"));
					break;

				case "DateTime":
					fieldXml.Add(new XAttribute("Format", field.Format ?? "DateOnly"));
					break;

				case "Number":
				case "Currency":
					if (field.MinValue.HasValue) fieldXml.Add(new XAttribute("Min", field.MinValue.Value));
					if (field.MaxValue.HasValue) fieldXml.Add(new XAttribute("Max", field.MaxValue.Value));
					if (field.Decimals.HasValue) fieldXml.Add(new XAttribute("Decimals", field.Decimals.Value));
					break;

				case "Boolean":
					// Default value for Boolean is usually 0 or 1
					if (!string.IsNullOrEmpty(field.DefaultValue))
					{
						fieldXml.Add(new XElement("Default", field.DefaultValue));
					}
					break;

				case "Calculated":
					if (!string.IsNullOrEmpty(field.Formula))
					{
						fieldXml.Add(new XAttribute("ResultType", field.ResultType ?? "Text"));
						fieldXml.Add(new XElement("Formula", field.Formula));
					}
                    if (!string.IsNullOrEmpty(field.Format))
                    {
                        fieldXml.Add(new XAttribute("Format", field.Format));
                    }
					

                    break;
                case "TargetTo":
                    // Поле нацеливания аудитории (Audience Targeting).
                    // В SharePoint XML оно всегда имеет тип "TargetTo".
                    // Обычно оно настраивается как многострочный текст (Note) внутри, 
                    // но в схеме создания указывается именно этот тип.
                    fieldXml.Attribute("Type").Value = "TargetTo";
                    // Добавляем технические атрибуты, характерные для этого поля
                    fieldXml.Add(new XAttribute("ReadOnly", "FALSE"));
                    fieldXml.Add(new XAttribute("Sortable", "FALSE"));
                    break;
            }

			
			return fieldXml.ToString();
		}
 
       public string GetAttributeFromXml(string xml, string attributeName)
        {
            if (string.IsNullOrEmpty(xml)) return null;
            try
            {
                var xDoc = System.Xml.Linq.XDocument.Parse(xml);
                return xDoc.Root?.Attribute(attributeName)?.Value;
            }
            catch
            {
                return null;
            }
        }
   
        public string CleanFieldXml(string rawXml)
		{
			bool compareField = false;
			return SPUtil.Infrastructure.SPUsingUtils.GetCleanFieldXml(rawXml, compareField);

		}
		public string CompareCleanFieldXml(string rawXml)
		{
			bool compareField = true;
			return SPUtil.Infrastructure.SPUsingUtils.GetCleanFieldXml(rawXml, compareField);

		}		
		
    }

}