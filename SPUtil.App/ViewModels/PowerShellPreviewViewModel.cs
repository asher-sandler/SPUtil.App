using Prism.Mvvm;
using SPUtil.Infrastructure;
using System.Collections.ObjectModel;
using System.Windows;

namespace SPUtil.App.ViewModels
{
    /// <summary>
    /// DataContext for UniversalPreviewWindow when showing a generated PowerShell script.
    /// </summary>
    public class PowerShellPreviewViewModel : BindableBase
    {
        private string _previewText    = string.Empty;
        private string _statusMessage  = string.Empty;
        private bool   _isExporting;
        private int    _exportProgress;
        private ObservableCollection<DialogButton> _dialogButtons = new();

        public string PreviewText    { get => _previewText;    set => SetProperty(ref _previewText,    value); }
        public string StatusMessage  { get => _statusMessage;  set => SetProperty(ref _statusMessage,  value); }
        public bool   IsExporting    { get => _isExporting;    set => SetProperty(ref _isExporting,    value); }
        public int    ExportProgress { get => _exportProgress; set => SetProperty(ref _exportProgress, value); }

        public ObservableCollection<DialogButton> DialogButtons
        {
            get => _dialogButtons;
            set => SetProperty(ref _dialogButtons, value);
        }

        public PowerShellPreviewViewModel(
            string script,
            string pageName,
            int webPartCount,
            Window ownerWindow,
            ObservableCollection<DialogButton> buttons)
        {
            PreviewText   = script;
            StatusMessage = $"Page: {pageName}  |  WebParts: {webPartCount}  |  " +
                            $"Script length: {script.Length} chars  |  " +
                            $"Click 'Copy script' to copy to clipboard";
            DialogButtons = buttons;
        }
    }
}
