using Prism.Mvvm;
using SPUtil.Infrastructure;
using System.Collections.ObjectModel;
using System.Windows;

namespace SPUtil.App.ViewModels
{
    /// <summary>
    /// DataContext for UniversalPreviewWindow when showing a page copy report.
    /// </summary>
    public class CopyReportPreviewViewModel : BindableBase
    {
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

        public CopyReportPreviewViewModel(string reportText, PageCopyReport report, Window ownerWindow)
        {
            PreviewText = reportText;

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
                    Caption  = "Close",
                    IsCancel = true,
                    Action   = () => ownerWindow?.Close()
                }
            };

            StatusMessage = report.Summary;
        }
    }
}