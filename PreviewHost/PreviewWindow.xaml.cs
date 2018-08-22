using System.Windows;
using System.Windows.Input;

namespace PreviewHost
{
    /// <summary>
    /// Interaction logic for PreviewWindow.xaml
    /// </summary>
    public partial class PreviewWindow : Window
    {
        string theFile;

        public string File
        {
            get
            {
                return theFile;
            }
            set
            {
                theFile = value;
                Title = theFile;
            }
        }

        public ICommand BrowseCommand { get; private set; }
        public ICommand RefreshCommand { get; private set; }

        public readonly DependencyProperty PreviewStatusTextProperty = DependencyProperty.Register("PreviewStatusText", typeof(string), typeof(PreviewWindow));

        public string PreviewStatusText
        {
            get { return (string)GetValue(PreviewStatusTextProperty); }
            set { SetValue(PreviewStatusTextProperty, value); }
        }

        public PreviewWindow(string file)
        {
            BrowseCommand = new BrowseCommand(this);
            RefreshCommand = new RefreshCommand(this);
            InitializeComponent();
            File = file;
            LoadPreview();
        }

        public void UnloadPreview()
        {
            // TODO: Unload the preview handler.
            PreviewStatusText = "Preview is not loaded.";
        }

        public void LoadPreview()
        {
            PreviewStatusText = "Preview is loading.";
            // TODO: Load the preview handler for File.
            // In case the handler crashes, this text is revealed to the user.
            // A real-world application might try restarting the handler before
            // giving up. Also, a read-world application should NOT put this
            // text visible to Narrator before the preview handler crashes.
            // For now, we just keep it simple.
            PreviewStatusText = "Preview handler crashed.";
        }

        private void PreviewContent_GotFocus(object sender, RoutedEventArgs e)
        {
            // TODO: Deal with focus changes.
        }
    }
}
