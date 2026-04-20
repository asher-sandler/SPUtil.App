using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Infrastructure;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Windows;

namespace SPUtil.App.ViewModels
{
    /// <summary>
    /// DataContext for UniversalPreviewWindow when showing WebPart properties.
    /// Owns its own PreviewText / DialogButtons / IsExporting so it is fully
    /// independent from MainWindowViewModel.
    /// </summary>
    public class WebPartsPreviewViewModel : BindableBase
    {
        // ── Properties required by UniversalPreviewWindow.xaml ───────────────
        private string _previewText = string.Empty;
        public string PreviewText
        {
            get => _previewText;
            set => SetProperty(ref _previewText, value);
        }

        private string _statusMessage = string.Empty;
        public string StatusMessage
        {
            get => _statusMessage;
            set => SetProperty(ref _statusMessage, value);
        }

        private bool _isExporting;
        public bool IsExporting
        {
            get => _isExporting;
            set => SetProperty(ref _isExporting, value);
        }

        private int _exportProgress;
        public int ExportProgress
        {
            get => _exportProgress;
            set => SetProperty(ref _exportProgress, value);
        }

        private ObservableCollection<DialogButton> _dialogButtons = new();
        public ObservableCollection<DialogButton> DialogButtons
        {
            get => _dialogButtons;
            set => SetProperty(ref _dialogButtons, value);
        }

        // ── Constructor — builds the text and wires the buttons ──────────────
        public WebPartsPreviewViewModel(
            IEnumerable<SPWebPartData> webParts,
            string pageTitle,
            Window ownerWindow)
        {
            // Build the formatted text block
            PreviewText = BuildPreviewText(webParts, pageTitle);

            // ── Buttons ──────────────────────────────────────────────────────
            DialogButtons = new ObservableCollection<DialogButton>
            {
                new DialogButton
                {
                    Caption = "📋  Copy all",
                    Action  = () =>
                    {
                        if (string.IsNullOrWhiteSpace(PreviewText)) return;
                        Clipboard.SetText(PreviewText);
                        StatusMessage = "✔ Copied to clipboard!";
                    }
                },

                new DialogButton
                {
                    Caption = "Close",
                    IsCancel = true,
                    Action  = () => ownerWindow?.Close()
                }
            };

            StatusMessage = $"Page: {pageTitle}  |  Web parts: {webParts.Count()}";
        }

        // ── Text builder ─────────────────────────────────────────────────────
        private static string BuildPreviewText(IEnumerable<SPWebPartData> webParts, string pageTitle)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"=== WebParts — {pageTitle} ===");
            sb.AppendLine($"Generated: {DateTime.Now:dd.MM.yyyy HH:mm:ss}");
            sb.AppendLine(new string('═', 60));

            int index = 1;
            foreach (var wp in webParts)
            {
                sb.AppendLine();
                sb.AppendLine($"[{index++}] {wp.Title}");
                sb.AppendLine($"    Type       : {wp.Type}");
                sb.AppendLine($"    ZoneId     : {wp.ZoneId}");

                // StorageKey is the GUID in the ms-rte-wpbox div — use this
                // to match this WebPart to the placeholder in PublishingPageContent
                sb.AppendLine($"    StorageKey : {wp.StorageKey}   ← matches div_{{GUID}} in page HTML");
                sb.AppendLine($"    Id         : {wp.Id}");

                if (wp.Properties != null && wp.Properties.Count > 0)
                {
                    sb.AppendLine("    ── Properties ──────────────────────────────");
                    foreach (var kv in wp.Properties.OrderBy(k => k.Key))
                    {
                        string val = kv.Value?.Replace("\n", "\n              ") ?? "";
                        sb.AppendLine($"    {kv.Key,-30} : {val}");
                    }
                }
                else
                {
                    sb.AppendLine("    (no properties)");
                }

                sb.AppendLine(new string('─', 60));
            }

            if (!webParts.Any())
                sb.AppendLine("(no web parts found for this page)");

            return sb.ToString();
        }
    }
}
