using Prism.Mvvm;
using SPUtil.Infrastructure;
using SPUtil.Services;
using System;
using System.Collections.ObjectModel;
using System.Windows;

namespace SPUtil.App.ViewModels
{
    public class PageCompareViewModel : BindableBase
    {
        private string _previewText    = string.Empty;
        private string _statusMessage  = string.Empty;
        private bool   _isExporting;
        private int    _exportProgress;
        private ObservableCollection<DialogButton> _dialogButtons = new();

        public string PreviewText   { get => _previewText;   set => SetProperty(ref _previewText,   value); }
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public bool   IsExporting   { get => _isExporting;   set => SetProperty(ref _isExporting,   value); }
        public int    ExportProgress { get => _exportProgress; set => SetProperty(ref _exportProgress, value); }
        public ObservableCollection<DialogButton> DialogButtons
        {
            get => _dialogButtons;
            set => SetProperty(ref _dialogButtons, value);
        }

        public PageCompareViewModel(
            PageCompareResult compareResult,
            string formattedText,
            Window ownerWindow,
            ISharePointService spService,
            Action onPlaceholdersInserted = null)
        {
            PreviewText   = formattedText;
            StatusMessage = compareResult.IsIdentical
                ? "✔ Pages are identical"
                : $"Differences:  ✏ {compareResult.ModifiedCount} modified  " +
                  $"➕ {compareResult.AddedCount} added  " +
                  $"➖ {compareResult.RemovedCount} removed";

            var buttons = new ObservableCollection<DialogButton>();

            // Copy
            buttons.Add(new DialogButton
            {
                Caption = "📋  Copy",
                Action  = () =>
                {
                    if (string.IsNullOrWhiteSpace(PreviewText)) return;
                    Clipboard.SetText(PreviewText);
                    StatusMessage = "✔ Copied to clipboard!";
                }
            });

            // Insert Placeholders — only when there are missing WebParts
            if (compareResult.RemovedCount > 0)
            {
                buttons.Add(new DialogButton
                {
                    Caption = $"🔧  Insert {compareResult.RemovedCount} Placeholder(s)",
                    Action  = async () =>
                    {
                        StatusMessage = "Inserting placeholders...";
                        IsExporting   = true;
                        try
                        {
                            await spService.InsertPlaceholdersAsync(
                                compareResult.TargetSiteUrl,
                                compareResult.TargetUrl,
                                compareResult);

                            StatusMessage =
                                $"✔ {compareResult.RemovedCount} placeholder(s) inserted on target. " +
                                $"Add the WebParts manually then run Sync Properties.";

                            onPlaceholdersInserted?.Invoke();
                        }
                        catch (Exception ex)
                        {
                            StatusMessage = $"✘ {ex.Message}";
                            MessageBox.Show($"Error inserting placeholders:\n{ex.Message}",
                                "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                        finally { IsExporting = false; }
                    }
                });
            }

            // Close
            buttons.Add(new DialogButton
            {
                Caption  = "Close",
                IsCancel = true,
                Action   = () => ownerWindow?.Close()
            });

            DialogButtons = buttons;
        }
    }
}
