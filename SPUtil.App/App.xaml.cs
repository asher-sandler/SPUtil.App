using Prism.DryIoc;
using Prism.Ioc;
using Prism.Mvvm;
using SPUtil.App.Views;
using SPUtil.App.ViewModels;
using SPUtil.Services;
using System.Windows;
using Serilog;

namespace SPUtil.App
{
    public partial class App : PrismApplication
    {
        protected override Window CreateShell()
        {
            return Container.Resolve<MainWindow>();
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            // Register service (Singleton — one instance for the whole application)
            containerRegistry.RegisterSingleton<ISharePointService, SharePointService>();
			containerRegistry.RegisterSingleton<SharePointService>(); 
			containerRegistry.RegisterSingleton<SharePointCloneService>();

            // Register view models in the container
            // We register them as plain types because we use them 
            // inside ContentControl via DataTemplate, not via RegionManager
            containerRegistry.Register<List100ViewModel>();
            containerRegistry.Register<Library101ViewModel>();

            // Register navigation (if you plan to use Journal/Navigate)
            containerRegistry.RegisterForNavigation<MainWindow, MainWindowViewModel>();
            containerRegistry.RegisterForNavigation<List100View, List100ViewModel>();
            containerRegistry.RegisterForNavigation<Library101View, Library101ViewModel>();

            ViewModelLocationProvider.Register<MainWindow, MainWindowViewModel>();
			
		// Register SharePointService both as interface and as concrete class
	        }
    }
}