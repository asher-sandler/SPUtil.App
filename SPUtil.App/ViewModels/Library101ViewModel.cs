using Prism.Commands;
using Prism.Mvvm;
using SPUtil.Infrastructure;
using SPUtil.Services;
using System;
using System.Collections.ObjectModel;
using System.Linq;
using System.Threading.Tasks;
using Serilog;

namespace SPUtil.App.ViewModels
{
    public class Library101ViewModel : BindableBase
    {
        private static readonly ILogger _log = Log.ForContext<Library101ViewModel>();

        private readonly ISharePointService _spService;
        private ObservableCollection<SPFileData> _files = new();
        private string _statusMessage = string.Empty;
        private string _libraryTitle = string.Empty;

        // Remembered from the initial load so NavigateToFolderAsync/NavigateUpAsync
        // can reissue GetLibraryItemsAsync without the caller passing them again.
        private string _siteUrl = string.Empty;
        private string _listId = string.Empty;

        // Root of the library — resolved once on the first load. Used to know
        // when "Up" has nowhere further to go (CanNavigateUp).
        private string _rootFolderPath = string.Empty;

        private string _currentFolderPath = string.Empty;
        public string CurrentFolderPath
        {
            get => _currentFolderPath;
            private set
            {
                if (SetProperty(ref _currentFolderPath, value))
                {
                    RaisePropertyChanged(nameof(CanNavigateUp));
                    RaisePropertyChanged(nameof(CurrentFolderDisplayPath));
                    RaisePropertyChanged(nameof(CurrentFolderUrl));
                }
            }
        }

        /// <summary>
        /// CurrentFolderPath with the site/library prefix stripped — shown in the
        /// UI instead of the full server-relative URL. E.g. "/home/huca/.../dev/
        /// LibraryTitle/פודר1" → "LibraryTitle/פודר1". Falls back to the raw
        /// path if, for some reason, it doesn't start with the resolved root
        /// (shouldn't normally happen, but better than showing an empty string).
        /// </summary>
        public string CurrentFolderDisplayPath
        {
            get
            {
                if (string.IsNullOrEmpty(_currentFolderPath)) return string.Empty;

                // Root's own parent — e.g. root "/home/huca/.../dev/LibraryTitle"
                // trimmed to its own last segment's parent, so the display keeps
                // the library name itself as the visible "top" instead of stripping
                // it away entirely (an empty display at the root reads as broken).
                string rootParent = ComputeParentPath(_rootFolderPath);

                if (_currentFolderPath.StartsWith(rootParent, StringComparison.OrdinalIgnoreCase))
                {
                    string tail = _currentFolderPath.Substring(rootParent.Length).TrimStart('/');
                    return string.IsNullOrEmpty(tail) ? "/" : tail;
                }

                return _currentFolderPath;
            }
        }

        /// <summary>
        /// Absolute URL to open the current folder in a browser — host from
        /// siteUrl + CurrentFolderPath (already server-relative from the domain
        /// root). Same pattern as DispFormUrl/PageUrl elsewhere in the project.
        /// Host is taken directly from siteUrl, not through NormalizeUrl — not
        /// verified through a load-balancer host (crs2/portals2/tss2), same
        /// caveat as those other two.
        /// </summary>
        public string CurrentFolderUrl
        {
            get
            {
                if (string.IsNullOrEmpty(_currentFolderPath) || string.IsNullOrEmpty(_siteUrl))
                    return string.Empty;

                try
                {
                    string hostRoot = "https://" + new Uri(_siteUrl).Host;
                    return $"{hostRoot}{_currentFolderPath}";
                }
                catch
                {
                    return string.Empty;
                }
            }
        }

        /// <summary>True once we are below the library root — i.e. "Up" has somewhere to go.</summary>
        public bool CanNavigateUp =>
            !string.IsNullOrEmpty(_currentFolderPath) &&
            !string.Equals(_currentFolderPath, _rootFolderPath, StringComparison.OrdinalIgnoreCase);

        public string LibraryTitle { get => _libraryTitle; set => SetProperty(ref _libraryTitle, value); }
        public string StatusMessage { get => _statusMessage; set => SetProperty(ref _statusMessage, value); }
        public ObservableCollection<SPFileData> Files { get => _files; set => SetProperty(ref _files, value); }
		// flag: this is the source pane
		private bool _isSourceMode;
		public bool IsSourceMode 
		{ 
			get => _isSourceMode; 
			set => SetProperty(ref _isSourceMode, value); 
		}

        private bool _isBusy;
        public bool IsBusy { get => _isBusy; set => SetProperty(ref _isBusy, value); }

        // Commands
        public DelegateCommand SelectAllCommand { get; }
        public DelegateCommand CopySelectedCommand { get; }
        public DelegateCommand DeleteSelectedCommand { get; }

        public Library101ViewModel(ISharePointService spService)
        {
            _spService = spService;

            // Initialize select-all command
            SelectAllCommand = new DelegateCommand(() =>
            {
                if (Files == null) return;
                foreach (var f in Files) f.IsSelected = true;
                
                // Re-assign collection to notify UI about internal list changes
                var temp = new ObservableCollection<SPFileData>(Files);
                Files = temp;
            });

            // Initialize copy command
            CopySelectedCommand = new DelegateCommand(() => {
                var selectedCount = Files?.Count(f => f.IsSelected) ?? 0;
                StatusMessage = $"STUB: Copying {selectedCount} item(s)...";
            });

            // Initialize delete command
            DeleteSelectedCommand = new DelegateCommand(() => {
                var selectedCount = Files?.Count(f => f.IsSelected) ?? 0;
                StatusMessage = $"STUB: Deleting {selectedCount} item(s)...";
            });
        }

        public async Task LoadDataAsync(string siteUrl, string listId)
        {
            IsBusy = true;
			const int maxRows = 500;
            await Task.Delay(300);
            try
            {
                StatusMessage = "Loading data from SharePoint...";
                string cleanId = listId.StartsWith("id:") ? listId.Substring(3) : listId;

                _siteUrl = siteUrl;
                _listId  = cleanId;

                // folderRelativeUrl omitted → service resolves the library root
                // and returns it as CurrentFolderPath, which we remember as the
                // "floor" for CanNavigateUp.
                var (data, resolvedPath) = await _spService.GetLibraryItemsAsync(siteUrl, cleanId);
                _rootFolderPath = resolvedPath;
                CurrentFolderPath = resolvedPath;

                ApplyLoadedItems(data, maxRows);
            }
            catch (Exception ex) 
            {
                _log.Caller().Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                StatusMessage = $"SERVER ERROR: {ex.Message}"; 
                System.Diagnostics.Debug.WriteLine($"Full error: {ex.ToString()}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        /// <summary>
        /// Loads the contents of a specific folder (drill-down on double-click).
        /// Reuses the siteUrl/listId remembered from LoadDataAsync — the caller
        /// only needs to pass the folder's server-relative URL (SPFileData.FullPath
        /// of the folder row that was double-clicked).
        /// </summary>
        public async Task NavigateToFolderAsync(string folderRelativeUrl)
        {
            IsBusy = true;
            const int maxRows = 500;
            try
            {
                StatusMessage = "Loading data from SharePoint...";

                var (data, resolvedPath) = await _spService.GetLibraryItemsAsync(_siteUrl, _listId, folderRelativeUrl);
                CurrentFolderPath = resolvedPath;

                ApplyLoadedItems(data, maxRows);
            }
            catch (Exception ex)
            {
                _log.Caller().Error(ex, "ERROR: {ExType} — {Message}", ex.GetType().Name, ex.Message);
                StatusMessage = $"SERVER ERROR: {ex.Message}";
                System.Diagnostics.Debug.WriteLine($"Full error: {ex.ToString()}");
            }
            finally
            {
                IsBusy = false;
            }
        }

        /// <summary>
        /// Navigates one level up from CurrentFolderPath. No-op if already at
        /// the library root (see CanNavigateUp) — callers should also disable
        /// the "Up" affordance in that state rather than relying solely on this guard.
        /// </summary>
        public async Task NavigateUpAsync()
        {
            if (!CanNavigateUp) return;

            string parent = ComputeParentPath(_currentFolderPath);
            await NavigateToFolderAsync(parent);
        }

        private void ApplyLoadedItems(List<SPFileData> data, int maxRows)
        {
            if (data.Count > maxRows)
            {
                StatusMessage = $"Warning: library contains {data.Count} items. Showing first {maxRows} only.";
                Files = new ObservableCollection<SPFileData>(data.Take(maxRows));
            }
            else
            {
                StatusMessage = $"Total items: {data.Count}";
                Files = new ObservableCollection<SPFileData>(data);
            }
        }

        /// <summary>Strips the last path segment of a server-relative URL ("/a/b/c" → "/a/b").</summary>
        private static string ComputeParentPath(string path)
        {
            string trimmed = path.TrimEnd('/');
            int idx = trimmed.LastIndexOf('/');
            return idx > 0 ? trimmed.Substring(0, idx) : trimmed;
        }
    }
}
