using System;
using System.Windows;
using System.Windows.Input;
using System.Windows.Media.Animation;
using System.Windows.Controls;


namespace SPUtil.App.Views
{
    public partial class AboutWindow : Window
    {
        public AboutWindow()
        {
            InitializeComponent();
        }

        // Open
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var openStory = (Storyboard)Resources["OpenAnimation"];
            openStory.Begin(this);

            var creditsStory = (Storyboard)Resources["CreditsAnimation"];
            creditsStory.Begin(this);
        }
		
		private void Email_Click(object sender, MouseButtonEventArgs e)
		{
			Clipboard.SetText(((TextBlock)sender).Text);
		}	

        // Animation on  OK
        private void BtnOk_Click(object sender, RoutedEventArgs e)
        {
            StartCloseAnimation();
        }
		

        // Move windows
        private void Border_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left)
            {
                DragMove();
            }
        }

        private void StartCloseAnimation()
        {
            var closeStory = (Storyboard)Resources["CloseAnimation"];
            closeStory.Begin(this);
        }

        // Close modal window
        private void CloseAnimation_Completed(object sender, EventArgs e)
        {
            Close();
        }
    }
}