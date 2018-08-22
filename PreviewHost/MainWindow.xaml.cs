using System.Windows;

namespace PreviewHost
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        bool isDropAllowed;

        private void Grid_DragEnter(object sender, DragEventArgs e)
        {
            e.Handled = true;
            isDropAllowed = e.Data.GetDataPresent(DataFormats.FileDrop, true);
            e.Effects = isDropAllowed ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void Grid_DragOver(object sender, DragEventArgs e)
        {
            e.Handled = true;
            e.Effects = isDropAllowed ? DragDropEffects.Copy : DragDropEffects.None;
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            e.Handled = true;
            if (!isDropAllowed)
                return;
            foreach (var file in (string[])e.Data.GetData(DataFormats.FileDrop, true))
                (new PreviewWindow(file)).Show();
        }
    }
}
