using System.IO;
using System.Windows;
using Microsoft.Win32;

namespace SPUtil.App.Views
{
    public partial class SchemaViewerWindow : Window
    {
        private readonly string _suggestedFileName;

        // Добавляем второй параметр в конструктор
        public SchemaViewerWindow(string xmlContent, string listName)
        {
            InitializeComponent();
            XmlBox.Text = xmlContent;

            // Формируем имя файла (например, "MyList_Schema.xml")
            _suggestedFileName = $"{listName}_Schema.xml";

            // Можно также поменять заголовок окна, чтобы было понятно, чья это схема
			

            this.Title = $"SP List '{listName}' Fields Schema";
        }

        private void Copy_Click(object sender, RoutedEventArgs e)
        {
            Clipboard.SetText(XmlBox.Text);
            MessageBox.Show("Schema was copied to clipboard!.", "Info");
        }

        private void Save_Click(object sender, RoutedEventArgs e)
        {
            // Используем _suggestedFileName вместо жестко прописанного "ListSchema.xml"
            SaveFileDialog sfd = new SaveFileDialog
            {
                Filter = "XML files (*.xml)|*.xml",
                FileName = _suggestedFileName
            };

            if (sfd.ShowDialog() == true)
            {
                File.WriteAllText(sfd.FileName, XmlBox.Text);
                MessageBox.Show("File was saved.", "Ok");
            }
        }
    }
}