using Prism.Mvvm;

namespace SPUtil.App.ViewModels
{
    public class InfoViewModel : BindableBase
    {
        private string _message = string.Empty;
        public string Message { get => _message; set => SetProperty(ref _message, value); }

        public InfoViewModel(string message)
        {
            Message = message;
        }
    }
}
