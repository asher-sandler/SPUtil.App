using Prism.DryIoc;
using Prism.Ioc;
using Prism.Mvvm;
using SPUtil.App.Views;
using SPUtil.App.ViewModels;
using SPUtil.Services;
using System;
using System.IO;
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

        protected override void OnStartup(StartupEventArgs e)
        {
            InitializeLogger();

            // Global handler — catches unhandled exceptions from async void methods
            // that would otherwise silently kill the process
            AppDomain.CurrentDomain.UnhandledException += (s, ev) =>
            {
                Log.Fatal(ev.ExceptionObject as Exception,
                    "Unhandled AppDomain exception — application will terminate");
                Log.CloseAndFlush();
            };

            // Catches exceptions on the UI thread
            DispatcherUnhandledException += (s, ev) =>
            {
                Log.Fatal(ev.Exception,
                    "Unhandled UI thread exception");
                ev.Handled = true;   // keep app alive, show the error
                MessageBox.Show(
                    $"Unexpected error:\n{ev.Exception.Message}\n\n" +
                    $"Details written to log file.",
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            };

            base.OnStartup(e);
            Log.Information("=== SPUtil started ===");
        }

        protected override void OnExit(ExitEventArgs e)
        {
            Log.Information("=== SPUtil exiting ===");
            Log.CloseAndFlush();   // flush all buffered log entries before exit
            base.OnExit(e);
        }

        protected override void RegisterTypes(IContainerRegistry containerRegistry)
        {
            containerRegistry.RegisterSingleton<ISharePointService, SharePointService>();
            containerRegistry.RegisterSingleton<SharePointService>();
            containerRegistry.RegisterSingleton<SharePointCloneService>();

            containerRegistry.Register<List100ViewModel>();
            containerRegistry.Register<Library101ViewModel>();

            containerRegistry.RegisterForNavigation<MainWindow, MainWindowViewModel>();
            containerRegistry.RegisterForNavigation<List100View, List100ViewModel>();
            containerRegistry.RegisterForNavigation<Library101View, Library101ViewModel>();

            ViewModelLocationProvider.Register<MainWindow, MainWindowViewModel>();
        }

        private static void InitializeLogger()
        {
            string logDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "SPUtil", "Logs");

            Directory.CreateDirectory(logDir);

            string logFile = Path.Combine(logDir, "sputil-.log");

            Log.Logger = new LoggerConfiguration()
                .MinimumLevel.Debug()
                .Enrich.FromLogContext()          // picks up Log.ForContext<T>() class names
                .WriteTo.File(
                    path:                   logFile,
                    rollingInterval:        RollingInterval.Day,   // new file each day
                    retainedFileCountLimit: 30,                    // keep 30 days
                    fileSizeLimitBytes:     20_000_000,            // 20 MB max per file
                    outputTemplate:
                        "{Timestamp:yyyy-MM-dd HH:mm:ss} [{Level:u3}] " +
                        "{SourceContext} — {Message:lj}{NewLine}" +
                        "{Exception}")
                .CreateLogger();

            Log.Information("Logger initialized. Log directory: {LogDir}", logDir);
        }
    }
}
