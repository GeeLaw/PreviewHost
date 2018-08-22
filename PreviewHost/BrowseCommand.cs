using Microsoft.Win32;
using System;
using System.Windows.Input;

namespace PreviewHost
{
    class BrowseCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;

        PreviewWindow previewWindow;

        public BrowseCommand(PreviewWindow owner)
        {
            if (owner == null)
                throw new ArgumentNullException("owner");
            previewWindow = owner;
        }

        public bool CanExecute(object parameter)
        {
            return true;
        }

        public void Execute(object parameter)
        {
            var ofd = new OpenFileDialog();
            ofd.Title = "Choose a file for preview";
            if (ofd.ShowDialog(previewWindow) != true)
                return;
            previewWindow.UnloadPreview();
            previewWindow.File = ofd.FileName;
            previewWindow.LoadPreview();
        }
    }
}
