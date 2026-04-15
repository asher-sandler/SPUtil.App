using Prism.DryIoc;
using Prism.Ioc;
using Prism.Mvvm;
using SPUtil.App.Views;
using SPUtil.App.ViewModels;
using SPUtil.Services;
using System.Windows;

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
            // Регистрация сервиса (Singleton - один экземпляр на всё приложение)
            containerRegistry.RegisterSingleton<ISharePointService, SharePointService>();
			containerRegistry.RegisterSingleton<SharePointService>(); 
			containerRegistry.RegisterSingleton<SharePointCloneService>();

            // Регистрация вью-моделей в контейнере
            // Мы регистрируем их просто как типы, так как мы используем их 
            // внутри ContentControl через DataTemplate, а не через RegionManager
            containerRegistry.Register<List100ViewModel>();
            containerRegistry.Register<Library101ViewModel>();

            // Регистрация навигации (если планируете использовать Journal/Navigate)
            containerRegistry.RegisterForNavigation<MainWindow, MainWindowViewModel>();
            containerRegistry.RegisterForNavigation<List100View, List100ViewModel>();
            containerRegistry.RegisterForNavigation<Library101View, Library101ViewModel>();

            ViewModelLocationProvider.Register<MainWindow, MainWindowViewModel>();
			
		// Регистрируем SharePointService и как интерфейс, и как сам класс
	        }
    }
}