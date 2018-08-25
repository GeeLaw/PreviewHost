using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Windows.Interop;
using System.Windows.Media;

namespace PreviewHost.Interop
{
    public sealed class PreviewHandler : IDisposable
    {
        #region IPreviewHandlerFrame support

        [StructLayout(LayoutKind.Sequential)]
        struct PreviewHandlerFrameInfo
        {
            public IntPtr AcceleratorTableHandle;
            public uint AcceleratorEntryCount;
        }

        [ComImport, Guid("fec87aaf-35f9-447a-adb7-20234491401a"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IPreviewHandlerFrame
        {
            [PreserveSig]
            HResult GetWindowContext(out PreviewHandlerFrameInfo pinfo);
            [PreserveSig]
            HResult TranslateAccelerator(ref MSG pmsg);
        }

        [ClassInterface(ClassInterfaceType.None)]
        sealed class PreviewHandlerFrame : IPreviewHandlerFrame, IDisposable
        {
            bool disposed;
            IPreviewHandlerManagedFrame site;

            public PreviewHandlerFrame(IPreviewHandlerManagedFrame frame)
            {
            }

            public void Dispose()
            {
            }

            public HResult GetWindowContext(out PreviewHandlerFrameInfo pinfo)
            {
                throw new NotImplementedException();
            }

            public HResult TranslateAccelerator(ref MSG pmsg)
            {
                throw new NotImplementedException();
            }
        }

        #endregion IPreviewHandlerFrame support

        #region IPreviewHandler major interfaces

        [ComImport, Guid("8895b1c6-b41f-4c1c-a562-0d564250836f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IPreviewHandler
        {
            [PreserveSig]
            HResult SetWindow(IntPtr hwnd, ref Rect prc);
            [PreserveSig]
            HResult SetRect(ref Rect prc);
            [PreserveSig]
            HResult DoPreview();
            [PreserveSig]
            HResult Unload();
            [PreserveSig]
            HResult SetFocus();
            [PreserveSig]
            HResult QueryFocus(out IntPtr phwnd);
            // TranslateAccelerator is not used here.
        }

        [ComImport, Guid("196bf9a5-b346-4ef0-aa1e-5dcdb76768b1"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IPreviewHandlerVisuals
        {
            [PreserveSig]
            HResult SetBackgroundColor(uint color);
            [PreserveSig]
            HResult SetFont(ref LogFontW plf);
            [PreserveSig]
            HResult SetTextColor(uint color);
        }

        static uint ColorRefFromColor(Color color)
        {
            return (((uint)color.B) << 16) | (((uint)color.G) << 8) | ((uint)color.R);
        }

        [ComImport, Guid("fc4801a3-2ba9-11cf-a229-00aa003d7352"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IObjectWithSite
        {
            [PreserveSig]
            HResult SetSite([In, MarshalAs(UnmanagedType.IUnknown)] object pUnkSite);
            // GetSite is not used.
        }

        #endregion IPreviewHandler major interfaces

        bool disposed;
        bool init;
        bool shown;
        PreviewHandlerFrame comSite;
        IPreviewHandlerManagedFrame site;
        IPreviewHandler previewHandler;
        IPreviewHandlerVisuals visuals;
        IntPtr pPreviewHandler;
        
        public PreviewHandler(Guid clsid, IPreviewHandlerManagedFrame frame)
        {
        }

        [Flags]
        enum ClassContext : uint
        {
            LocalServer = 0x4
        }

        [DllImport("ole32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern HResult CoCreateInstance(ref Guid rclsid, IntPtr pUnkOuter, ClassContext dwClsContext, ref Guid riid, out IntPtr ppv);

        #region Initialization interfaces

        [ComImport, Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IInitializeWithStream
        {
            [PreserveSig]
            HResult Initialize(IStream psi, StorageMode grfMode);
        }

        [ComImport, Guid("b824b49d-22ac-4161-ac8a-9916e8fa3f7f"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IInitializeWithStreamNative
        {
            [PreserveSig]
            HResult Initialize(IntPtr psi, StorageMode grfMode);
        }

        static readonly Guid IInitializeWithStreamIid = Guid.ParseExact("b824b49d-22ac-4161-ac8a-9916e8fa3f7f", "d");

        [ComImport, Guid("b7d14566-0509-4cce-a71f-0a554233bd9b"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IInitializeWithFile
        {
            [PreserveSig]
            HResult Initialize([In] string pszFilePath, StorageMode grfMode);
        }

        static readonly Guid IInitializeWithFileIid = Guid.ParseExact("b7d14566-0509-4cce-a71f-0a554233bd9b", "d");

        [ComImport, Guid("7f73be3f-fb79-493c-a6c7-7ee14e245841"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IInitializeWithItem
        {
            [PreserveSig]
            HResult Initialize(IntPtr psi, StorageMode grfMode);
        }

        static readonly Guid IInitializeWithItemIid = Guid.ParseExact("7f73be3f-fb79-493c-a6c7-7ee14e245841", "d");

        #endregion
        
        #region IDisposable pattern

        void Dispose(bool disposing)
        {
        }

        ~PreviewHandler()
        {
            Dispose(false);
        }

        void IDisposable.Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        
        #endregion

    }
}
