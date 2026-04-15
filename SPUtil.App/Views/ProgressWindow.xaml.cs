using System;
using System.Diagnostics;
using System.Threading;
using System.Windows;

namespace SPUtil.Views
{
    public partial class ProgressWindow : Window
    {
        // Поле должно быть доступно для метода отмены
        private readonly CancellationTokenSource _cts;
        private readonly Stopwatch _stopwatch;

        // Конструктор ТЕПЕРЬ принимает 1 аргумент, как и просит ViewModel
        public ProgressWindow(CancellationTokenSource cts)
        {
            InitializeComponent();
            _cts = cts;

            _stopwatch = new Stopwatch();
            _stopwatch.Start();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            // Теперь это будет реально останавливать поток в SharePointService
            var result = MessageBox.Show("Are you sure you want to cancel?",
                "Cancel", MessageBoxButton.YesNo, MessageBoxImage.Question);

            if (result == MessageBoxResult.Yes)
            {
                _cts?.Cancel(); // Посылаем сигнал отмены
                this.Close();
            }
        }

        public void UpdateStatus(int processed, int total, string message)
        {
            Dispatcher.Invoke(() => {
                TxtStatus.Text = message;

                if (total > 0)
                {
                    double percent = (double)processed / total * 100;
                    PrgBar.Value = percent;
                    TxtCount.Text = $"Processed: {processed} / {total} ({percent:F1}%)";

                    // Расчет времени
                    if (processed > 0)
                    {
                        long elapsedMs = _stopwatch.ElapsedMilliseconds;
                        long avgMsPerItem = elapsedMs / processed;
                        long remainingItems = total - processed;
                        TimeSpan t = TimeSpan.FromMilliseconds(remainingItems * avgMsPerItem);
                        TxtTime.Text = string.Format("Remaining: {0:D2}m:{1:D2}s", t.Minutes, t.Seconds);
                    }
                }
            });
        }
    }
}