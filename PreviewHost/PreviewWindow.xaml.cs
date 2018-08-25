using PreviewHost.Interop;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using System.Windows.Media;
using System.Windows.Threading;

namespace PreviewHost
{
    /// <summary>
    /// Interaction logic for PreviewWindow.xaml
    /// </summary>
    public partial class PreviewWindow : Window, IPreviewHandlerManagedFrame
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
        public ICommand ExitCommand { get; private set; }

        public static readonly DependencyProperty PreviewStatusTextProperty = DependencyProperty.Register("PreviewStatusText", typeof(string), typeof(PreviewWindow));

        public string PreviewStatusText
        {
            get { return (string)GetValue(PreviewStatusTextProperty); }
            set { SetValue(PreviewStatusTextProperty, value); }
        }

        WindowInteropHelper interop;
        PreviewHandler currentHandler;

        public PreviewWindow(string file)
        {
            BrowseCommand = new BrowseCommand(this);
            RefreshCommand = new RefreshCommand(this);
            ExitCommand = new ExitCommand(this);
            InitializeComponent();
            File = file;
            interop = new WindowInteropHelper(this);
            interop.EnsureHandle();
        }

        public void UnloadPreview()
        {
            PreviewStatusText = "Preview is not loaded.";
            if (currentHandler == null)
                return;
            currentHandler.UnloadPreview();
            currentHandler = null;
        }

        public void LoadPreview()
        {
            if (currentHandler != null)
                return;
            PreviewStatusText = "Preview is loading.";
            var clsid = Interop.PreviewHandlerDiscovery.FindPreviewHandlerFor(Path.GetExtension(File), interop.Handle);
            if (clsid == null)
            {
                PreviewStatusText = "No preview handler is associated with this file type.";
                return;
            }
            IntPtr pobj = IntPtr.Zero;
            try
            {
                currentHandler = new PreviewHandler(clsid.Value, this);
                currentHandler.InitWithFileWithEveryWay(File);
            }
            catch (Exception ex)
            {
                UnloadPreview();
                PreviewStatusText = "Preview handler is malfunctioning:\r\n" + ex.Message;
                return;
            }
            currentHandler.SetBackground(((SolidColorBrush)Background).Color);
            currentHandler.SetForeground(((SolidColorBrush)Foreground).Color);
            currentHandler.DoPreview();
            // In case the handler crashes, this text is revealed to the user.
            // A real-world application might try restarting the handler before
            // giving up. Also, a read-world application should NOT put this
            // text visible to Narrator before the preview handler crashes.
            // For now, we just keep it simple.
            PreviewStatusText = "A preview should be here.";
        }

        private void PreviewContent_GotFocus(object sender, RoutedEventArgs e)
        {
            if (currentHandler != null)
            {
                currentHandler.Focus();
                if (currentHandler.QueryFocus() == IntPtr.Zero)
                {
                    var old = currentHandler;
                    currentHandler = null;
                    contentPresenter.Focus();
                    currentHandler = old;
                }
            }
        }

        public IntPtr WindowHandle => interop.Handle;

        public Interop.Rect PreviewerBounds
        {
            get
            {
                var source = PresentationSource.FromVisual(contentPresenter);
                var transformToDevice = source.CompositionTarget.TransformToDevice;
                var transformToWindow = contentPresenter.TransformToAncestor(this);
                var physicalSize = (Size)transformToDevice.Transform((Vector)contentPresenter.RenderSize);
                var physicalPos = transformToDevice.Transform(transformToWindow.Transform(new Point(0, 0)));
                Interop.Rect result = new Interop.Rect();
                result.Left = (int)(physicalPos.X + 0.5);
                result.Top = (int)(physicalPos.Y + 0.5);
                result.Right = (int)(physicalPos.X + physicalSize.Width + 0.5);
                result.Bottom = (int)(physicalPos.Y + physicalSize.Height + 0.5);
                return result;
            }
        }

        // We will be interested in Tab, Ctrl+B, F5 and Esc.
        public IEnumerable<AcceleratorEntry> GetAcceleratorTable()
        {
            yield return new AcceleratorEntry { IsVirtual = 1, Key = 0x09 }; // Tab
            yield return new AcceleratorEntry { IsVirtual = 1, Key = 0x42 }; // B
            yield return new AcceleratorEntry { IsVirtual = 1, Key = 0x74 }; // F5
            yield return new AcceleratorEntry { IsVirtual = 1, Key = 0x1b }; // Esc
        }

        public bool TranslateAccelerator(ref MSG msg)
        {
            if (msg.message != 0x100)
                return false;
            var vk = msg.wParam.ToInt64();
            ModifierKeys modState = ModifierKeys.None;
            Dispatcher.Invoke(() => modState = Keyboard.Modifiers);
            Action action = null;
            try
            {
                switch (vk)
                {
                    case 0x09:
                        if ((modState & ModifierKeys.Shift) == ModifierKeys.Shift)
                        {
                            action = () => contentPresenter.MoveFocus(new TraversalRequest(FocusNavigationDirection.Previous));
                        }
                        else
                        {
                            action = () => contentPresenter.MoveFocus(new TraversalRequest(FocusNavigationDirection.Next));
                        }
                        return true;
                    case 0x42:
                        if ((modState & (ModifierKeys.Control | ModifierKeys.Alt | ModifierKeys.Shift)) == ModifierKeys.Control)
                        {
                            action = () =>
                            {
                                BrowseCommand.Execute(null);
                                contentPresenter.Focus();
                            };
                            return true;
                        }
                        return false;
                    case 0x74:
                        if ((modState & (ModifierKeys.Control | ModifierKeys.Alt | ModifierKeys.Shift | ModifierKeys.Windows)) == ModifierKeys.None)
                        {
                            action = () =>
                            {
                                RefreshCommand.Execute(null);
                                contentPresenter.Focus();
                            };
                            return true;
                        }
                        return false;
                    case 0x1b:
                        if ((modState & (ModifierKeys.Control | ModifierKeys.Alt | ModifierKeys.Shift | ModifierKeys.Windows)) == ModifierKeys.None)
                        {
                            action = () => ExitCommand.Execute(null);
                            return true;
                        }
                        return false;
                    default:
                        return false;
                }
            }
            finally
            {
                if (action != null)
                    Dispatcher.InvokeAsync(action, DispatcherPriority.Send);
            }
        }

        private void PreviewHost_Closing(object sender, CancelEventArgs e)
        {
            UnloadPreview();
        }

        private void PreviewHost_Loaded(object sender, RoutedEventArgs e)
        {
            LoadPreview();
        }

        private void PreviewHost_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (currentHandler != null)
                currentHandler.ResetBounds();
        }
    }
}
