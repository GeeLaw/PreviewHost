using System;
using System.Windows.Input;

namespace PreviewHost
{
    class ExitCommand : ICommand
    {
        public event EventHandler CanExecuteChanged;

        PreviewWindow previewWindow;

        public ExitCommand(PreviewWindow owner)
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
            previewWindow.Close();
        }
    }
}
