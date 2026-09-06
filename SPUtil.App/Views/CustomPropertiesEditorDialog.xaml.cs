using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using SPUtil.Infrastructure;

namespace SPUtil.Views
{
    // Shows editable controls for a curated subset of custom properties
    // (path/file-bound text values, boolean flags, and the two known global
    // enums). The dialog itself never touches ISharePointService — it only
    // collects edits into EditedProperties and returns them via DialogResult,
    // same convention as every other dialog in this app (e.g. ExistsActionDialog).
    // The caller (PagesViewModel.ExecuteEditCustomProperties) performs the
    // actual save via SaveSingleWebPartPropertiesAsync and reports the result.
    //
    // TODO / IDEA (raised 06.09.2026, needs discussion with AI before implementing):
    // Add an EKMD | ADA switch at the top of this form. When ADA is selected and a
    // property value contains the substring "\\ekeksql00\sp_resources$" (an EKMD-only
    // file-server UNC path), replace that substring across ALL fields in the form with
    // "\\vscifs.cc.huji.ac.il\adadata\SP_Resources" — the equivalent ADA-side path.
    // Not implemented yet — needs a decision on: where the switch's initial value
    // comes from (target site URL domain?), whether the replacement is case-sensitive,
    // and whether it should apply live as the user types or only as a one-time
    // "Convert to ADA" action.
    public partial class CustomPropertiesEditorDialog : Window
    {
        // Same base skip-list as ParseExportXmlProperties (PagesViewModel.cs) — standard
        // WebPart chrome/base properties are never candidates for this editor.
        private static readonly HashSet<string> _baseSkip = new(StringComparer.OrdinalIgnoreCase)
        {
            "AllowClose","AllowConnect","AllowEdit","AllowHide","AllowMinimize",
            "AllowZoneChange","AuthorizationFilter","CatalogIconImageUrl",
            "ChromeState","ChromeType","Direction","ExportMode","HelpMode",
            "HelpUrl","Hidden","ImportErrorMessage","TitleIconImageUrl","TitleUrl",
            "Title","Description"
        };

        // Properties that match the "path/file" name heuristic but were manually
        // reviewed and rejected as not user-editable in this form:
        //  - controlPath/_dataPath/_configPath: structural XML-tag paths inside the
        //    control's own config format, not real file/network paths.
        //  - CSS_Path/CssPath/JsPath: static site-relative style/script references,
        //    not meant to be edited per-page.
        //  - loadFormByIdPageURL/successPageUrl: page-flow URLs reviewed and excluded
        //    by the user (06.09.2026).
        //  - initiatorMailTemplate/approverMailTemplate: reviewed and excluded
        //    by the user (06.09.2026).
        private static readonly HashSet<string> _pathFileExclusions = new(StringComparer.OrdinalIgnoreCase)
        {
            "controlPath", "_dataPath", "_configPath",
            "CSS_Path", "CssPath", "JsPath",
            "loadFormByIdPageURL", "successPageUrl",
            "initiatorMailTemplate", "approverMailTemplate",
			"CSS_AUX_Path","CSS_Style_Path","Css_File_Name",
			"Css_Path","Css_File_Folder"
        };

        // Properties that do NOT match the path/file name heuristic but were
        // explicitly asked for anyway (06.09.2026) — shown in the same
        // "Paths & Files" list; boolean values still render as a CheckBox there.
        private static readonly HashSet<string> _explicitIncludes = new(StringComparer.OrdinalIgnoreCase)
        {
            "showLog","DebugMode","adminMode",
			"Debug","Show_Debug","ShowDebugLog"
        };

        private readonly ObservableCollection<PropertyItem> _pathFileItems = new();
        private readonly ObservableCollection<EnumPropertyItem> _enumItems = new();

        /// <summary>
        /// Populated by BtnSave_Click. Empty when the dialog is cancelled or when
        /// there was nothing editable to show in the first place.
        /// </summary>
        public Dictionary<string, string> EditedProperties { get; private set; } = new();

        public CustomPropertiesEditorDialog(SPWebPartData webPart)
        {
            InitializeComponent();

            TxtWpTitle.Text = $"WebPart: {webPart.Title}";
            TxtWpStorageKey.Text = $"StorageKey: {webPart.StorageKey}";

            foreach (var kv in webPart.Properties)
            {
                if (_baseSkip.Contains(kv.Key))
                    continue;

                if (WebPartEnumHelper.IsKnownEnum(kv.Key))
                {
                    WebPartEnumHelper.TryGetLabel(kv.Key, kv.Value, out string label);
                    _enumItems.Add(new EnumPropertyItem
                    {
                        Name = kv.Key,
                        Options = WebPartEnumHelper.GetOptions(kv.Key).ToList(),
                        SelectedValue = label
                    });
                    continue;
                }

                bool include = (LooksLikePathOrFile(kv.Key) && !_pathFileExclusions.Contains(kv.Key))
                               || _explicitIncludes.Contains(kv.Key);

                if (include)
                {
                    _pathFileItems.Add(PropertyItem.Create(kv.Key, kv.Value));
                }
            }

            PathFileList.ItemsSource = _pathFileItems;
            EnumList.ItemsSource = _enumItems;

            PathFileSection.Visibility = _pathFileItems.Any() ? Visibility.Visible : Visibility.Collapsed;
            EnumSection.Visibility     = _enumItems.Any()     ? Visibility.Visible : Visibility.Collapsed;
            TxtNoProps.Visibility      = (!_pathFileItems.Any() && !_enumItems.Any())
                                            ? Visibility.Visible : Visibility.Collapsed;

            // Nothing to edit — Save would be a no-op, so disable it rather than
            // let the user click through to an empty result.
            BtnSave.IsEnabled = _pathFileItems.Any() || _enumItems.Any();
        }

        private static bool LooksLikePathOrFile(string propertyName)
        {
            string n = propertyName.ToLowerInvariant();
            return n.Contains("path") || n.Contains("file") || n.Contains("url") || n.Contains("template");
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var edited = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            // Paths & Files — text and boolean rows both live in the same list;
            // read whichever value is the "real" one for that row.
            foreach (var item in _pathFileItems)
            {
                edited[item.Name] = item.IsBoolean
                    ? item.BoolValue.ToString()
                    : item.Value;
            }

            // Text Layout — convert the selected label ("right"/"rtl"...) back to
            // the raw "0"/"1" value the server actually stores.
            foreach (var item in _enumItems)
            {
                if (WebPartEnumHelper.TryGetRawValue(item.Name, item.SelectedValue, out string rawValue))
                    edited[item.Name] = rawValue;
            }

            EditedProperties = edited;
            DialogResult = true;
            Close();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }

        // Plain POCO for ItemsControl binding — no INotifyPropertyChanged needed;
        // values are read once, on Save, not observed live elsewhere.
        // Value-based boolean detection: some properties matched by the path/file name
        // heuristic (e.g. "Load_Form_With_No_Parameters_In_URL") actually hold True/False,
        // not a real path — those render as a CheckBox instead of a TextBox.
        private class PropertyItem
        {
            public string Name  { get; set; } = string.Empty;
            public string Value { get; set; } = string.Empty;
            public bool IsBoolean { get; set; }
            public bool IsText => !IsBoolean;
            public bool BoolValue { get; set; }

            public static PropertyItem Create(string name, string rawValue)
            {
                bool isBool = bool.TryParse(rawValue, out bool parsed);
                return new PropertyItem
                {
                    Name = name,
                    Value = rawValue,
                    IsBoolean = isBool,
                    BoolValue = isBool && parsed
                };
            }
        }

        private class EnumPropertyItem
        {
            public string Name { get; set; } = string.Empty;
            public List<string> Options { get; set; } = new();
            public string SelectedValue { get; set; } = string.Empty;
        }
    }
}
