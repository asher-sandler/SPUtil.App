using System.Collections.ObjectModel;

namespace SPUtil.Infrastructure
{
    public class SPNode
    {
        public string Title { get; set; } = string.Empty;
        public string Path { get; set; } = string.Empty;
        public SharePointObjectType Type { get; set; }
        public string? Tag { get; set; } 

        public ObservableCollection<SPNode> Children { get; set; } = new ObservableCollection<SPNode>();

        public SPNode() { }

        public SPNode(string title, string path, SharePointObjectType type)
        {
            Title = title;
            Path = path;
            Type = type;
        }
    }

    public enum SharePointObjectType
    {
        Site,
        List,
        Folder,
        File
    }
}