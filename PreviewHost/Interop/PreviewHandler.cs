using System;
using System.Collections.Generic;
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
                disposed = true;
                if (frame == null)
                    throw new ArgumentNullException("frame");
                disposed = false;
                site = frame;
            }

            public void Dispose()
            {
                disposed = true;
                site = null;
            }

            public HResult GetWindowContext(out PreviewHandlerFrameInfo pinfo)
            {
                pinfo.AcceleratorTableHandle = IntPtr.Zero;
                pinfo.AcceleratorEntryCount = 0;
                if (disposed)
                    return HResult.E_FAIL;
                List<AcceleratorEntry> list;
                try
                {
                    list = new List<AcceleratorEntry>(32767);
                    foreach (var accel in site.GetAcceleratorTable())
                    {
                        if (list.Count == 32767)
                        {
                            list.Clear();
                            break;
                        }
                        list.Add(accel);
                    }
                    if (list.Count == 0)
                        return HResult.E_OUTOFMEMORY;
                    var arr = list.ToArray();
                    var hAccel = AcceleratorEntry.CreateAcceleratorTable(arr, arr.Length);
                    if (hAccel == IntPtr.Zero)
                        return HResult.E_OUTOFMEMORY;
                    pinfo.AcceleratorTableHandle = hAccel;
                    pinfo.AcceleratorEntryCount = (uint)arr.Length;
                    return HResult.S_OK;
                }
                catch (OutOfMemoryException)
                {
                    return HResult.E_OUTOFMEMORY;
                }
                catch (Exception e)
                {
                    Environment.FailFast("IPreviewHandlerFrame.GetWindowContext failed.", e);
                    // Unreachable.
                    return HResult.E_FAIL;
                }
            }

            public HResult TranslateAccelerator(ref MSG pmsg)
            {
                if (disposed)
                    return HResult.E_FAIL;
                try
                {
                    return site.TranslateAccelerator(ref pmsg) ? HResult.S_OK : HResult.S_FALSE;
                }
                catch (OutOfMemoryException)
                {
                    return HResult.E_OUTOFMEMORY;
                }
                catch (Exception e)
                {
                    Environment.FailFast("IPreviewHandlerFrame.TranslateAccelerator failed.", e);
                    // Unreachable.
                    return HResult.E_FAIL;
                }
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
            disposed = true;
            init = false;
            shown = false;
            comSite = new PreviewHandlerFrame(frame);
            site = frame;
            try
            {
                SetupHandler(clsid);
                disposed = false;
            }
            catch
            {
                if (previewHandler != null)
                {
                    Marshal.ReleaseComObject(previewHandler);
                    previewHandler = null;
                }
                Marshal.Release(pPreviewHandler);
                pPreviewHandler = IntPtr.Zero;
                comSite.Dispose();
                comSite = null;
                site = null;
                throw;
            }
        }

        [Flags]
        enum ClassContext : uint
        {
            LocalServer = 0x4
        }

        [DllImport("ole32.dll", CallingConvention = CallingConvention.StdCall)]
        static extern HResult CoCreateInstance(ref Guid rclsid, IntPtr pUnkOuter, ClassContext dwClsContext, ref Guid riid, out IntPtr ppv);

        static readonly Guid IPreviewHandlerIid = Guid.ParseExact("8895b1c6-b41f-4c1c-a562-0d564250836f", "d");

        void SetupHandler(Guid clsid)
        {
            IntPtr pph;
            var iid = IPreviewHandlerIid;
            var cannotCreate = "Cannot create class " + clsid.ToString() + " as IPreviewHandler.";
            var cannotCast = "Cannot cast class " + clsid.ToString() + " as IObjectWithSite.";
            // Important: manully calling CoCreateInstance is necessary.
            // If we use Activator.CreateInstance(Type.GetTypeFromCLSID(...)),
            // CLR will allow in-process server, which defeats isolation and
            // creates strange bugs.
            HResult hr = CoCreateInstance(ref clsid, IntPtr.Zero, ClassContext.LocalServer, ref iid, out pph);
            // See https://blogs.msdn.microsoft.com/adioltean/2005/06/24/when-cocreateinstance-returns-0x80080005-co_e_server_exec_failure/
            // CO_E_SERVER_EXEC_FAILURE also tends to happen when debugging in Visual Studio.
            // Moreover, to create the instance in a server at low integrity level, we need
            // to use another thread with low mandatory label. We keep it simple by creating
            // a same-integrity object.
            if (hr == HResult.CO_E_SERVER_EXEC_FAILURE)
                hr = CoCreateInstance(ref clsid, IntPtr.Zero, ClassContext.LocalServer, ref iid, out pph);
            if ((int)hr < 0)
                throw new COMException(cannotCreate, (int)hr);
            pPreviewHandler = pph;
            previewHandler = (IPreviewHandler)Marshal.GetUniqueObjectForIUnknown(pph);
            var objectWithSite = (IObjectWithSite)previewHandler;
            hr = objectWithSite.SetSite(comSite);
            if ((int)hr < 0)
                throw new COMException("Cannot set site to the preview handler object.", (int)hr);
            visuals = previewHandler as IPreviewHandlerVisuals;
        }

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
            HResult Initialize([MarshalAs(UnmanagedType.LPWStr)] string pszFilePath, StorageMode grfMode);
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

        /// <summary>
        /// Tries to initialize the preview handler with an IStream.
        /// </summary>
        /// <exception cref="COMException">This exception is thrown if QueryInterface fails for reason other than E_NOINTERFACE, or if IInitializeWithStream.Initialize fails for reason other than E_NOTIMPL.</exception>
        /// <exception cref="ArgumentOutOfRangeException">Thrown if mode is neither Read nor ReadWrite.</exception>
        /// <param name="stream">The IStream interface used to initialize the preview handler.</param>
        /// <param name="mode">The storage mode, must be Read or ReadWrite.</param>
        /// <returns>If the handler supports initialization with IStream, true; otherwise, false.</returns>
        public bool InitWithStream(IStream stream, StorageMode mode)
        {
            if (mode != StorageMode.Read && mode != StorageMode.ReadWrite)
                throw new ArgumentOutOfRangeException("mode", mode, "The argument mode must be Read or ReadWrite.");
            var iid = IInitializeWithStreamIid;
            IntPtr piws;
            var hr = Marshal.QueryInterface(pPreviewHandler, ref iid, out piws);
            if (hr == (int)HResult.E_NOINTERFACE || piws == IntPtr.Zero)
                return false;
            var iws = (IInitializeWithStream)Marshal.GetUniqueObjectForIUnknown(piws);
            try
            {
                var hrr = iws.Initialize(stream, mode);
                if (hrr == HResult.E_NOTIMPL)
                    return false;
                if ((int)hrr < 0)
                    throw new COMException("IInitializeWithStream.Initialize failed.", (int)hrr);
                init = true;
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(iws);
            }
        }

        /// <summary>
        /// Same as InitWithStream(IStream, StorageMode).
        /// </summary>
        /// <exception cref="COMException">See InitWithStream(IStream, StorageMode).</exception>
        /// <exception cref="ArgumentOutOfRangeException">See InitWithStream(IStream, StorageMode).</exception>
        /// <param name="pStream">The native pointer to the IStream interface.</param>
        /// <param name="mode">The storage mode.</param>
        /// <returns>True or false, see InitWithStream(IStream, StorageMode).</returns>
        public bool InitWithStream(IntPtr pStream, StorageMode mode)
        {
            EnsureNotDisposed();
            EnsureNotInitialized();
            if (mode != StorageMode.Read && mode != StorageMode.ReadWrite)
                throw new ArgumentOutOfRangeException("mode", mode, "The argument mode must be Read or ReadWrite.");
            var iid = IInitializeWithStreamIid;
            IntPtr piws;
            var hr = Marshal.QueryInterface(pPreviewHandler, ref iid, out piws);
            if (hr == (int)HResult.E_NOINTERFACE || piws == IntPtr.Zero)
                return false;
            var iws = (IInitializeWithStreamNative)Marshal.GetUniqueObjectForIUnknown(piws);
            try
            {
                var hrr = iws.Initialize(pStream, mode);
                if (hrr == HResult.E_NOTIMPL)
                    return false;
                if ((int)hrr < 0)
                    throw new COMException("IInitializeWithStream.Initialize failed.", (int)hrr);
                init = true;
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(iws);
            }
        }

        /// <summary>
        /// Same as InitWithStream(IStream, StorageMode).
        /// </summary>
        /// <exception cref="COMException">See InitWithStream(IStream, StorageMode).</exception>
        /// <exception cref="ArgumentOutOfRangeException">See InitWithStream(IStream, StorageMode).</exception>
        /// <param name="psi">The native pointer to the IShellItem interface.</param>
        /// <param name="mode">The storage mode.</param>
        /// <returns>True or false, see InitWithStream(IStream, StorageMode).</returns>
        public bool InitWithItem(IntPtr psi, StorageMode mode)
        {
            EnsureNotDisposed();
            EnsureNotInitialized();
            if (mode != StorageMode.Read && mode != StorageMode.ReadWrite)
                throw new ArgumentOutOfRangeException("mode", mode, "The argument mode must be Read or ReadWrite.");
            var iid = IInitializeWithItemIid;
            IntPtr piwi;
            var hr = Marshal.QueryInterface(pPreviewHandler, ref iid, out piwi);
            if (hr == (int)HResult.E_NOINTERFACE || piwi == IntPtr.Zero)
                return false;
            var iwi = (IInitializeWithItem)Marshal.GetUniqueObjectForIUnknown(piwi);
            try
            {
                var hrr = iwi.Initialize(psi, mode);
                if (hrr == HResult.E_NOTIMPL)
                    return false;
                if ((int)hrr < 0)
                    throw new COMException("IInitializeWithItem.Initialize failed.", (int)hrr);
                init = true;
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(iwi);
            }
        }

        /// <summary>
        /// Same as InitWithStream(IStream, StorageMode).
        /// </summary>
        /// <exception cref="COMException">See InitWithStream(IStream, StorageMode).</exception>
        /// <exception cref="ArgumentOutOfRangeException">See InitWithStream(IStream, StorageMode).</exception>
        /// <param name="path">The path to the file.</param>
        /// <param name="mode">The storage mode.</param>
        /// <returns>True or false, see InitWithStream(IStream, StorageMode).</returns>
        public bool InitWithFile(string path, StorageMode mode)
        {
            EnsureNotDisposed();
            EnsureNotInitialized();
            if (mode != StorageMode.Read && mode != StorageMode.ReadWrite)
                throw new ArgumentOutOfRangeException("mode", mode, "The argument mode must be Read or ReadWrite.");
            var iid = IInitializeWithFileIid;
            IntPtr piwf;
            var hr = Marshal.QueryInterface(pPreviewHandler, ref iid, out piwf);
            if (hr == (int)HResult.E_NOINTERFACE || piwf == IntPtr.Zero)
                return false;
            var iwf = (IInitializeWithFile)Marshal.GetUniqueObjectForIUnknown(piwf);
            try
            {
                var hrr = iwf.Initialize(path, mode);
                if (hrr == HResult.E_NOTIMPL)
                    return false;
                if ((int)hrr < 0)
                    throw new COMException("IInitializeWithFile.Initialize failed.", (int)hrr);
                init = true;
                return true;
            }
            finally
            {
                Marshal.ReleaseComObject(iwf);
            }
        }

        /// <summary>
        /// Tries each way to initialize the object with a file.
        /// </summary>
        /// <param name="path">The file name.</param>
        /// <returns>If initialization was successful, true; otherwise, an exception is thrown.</returns>
        public bool InitWithFileWithEveryWay(string path)
        {
            var exceptions = new List<Exception>();
            var pobj = IntPtr.Zero;
            // Why should we try IStream first?
            // Because that gives us the best security.
            // If we initialize with string or IShellItem,
            // we have no control over how the preview handler
            // opens the file, which might decide to open the
            // file for read/write exclusively.
            try
            {
                pobj = ItemStreamHelper.IStreamFromPath(path);
                if (pobj != IntPtr.Zero
                    && InitWithStream(pobj, StorageMode.Read))
                    return true;
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }
            finally
            {
                if (pobj != IntPtr.Zero)
                    ItemStreamHelper.ReleaseObject(pobj);
                pobj = IntPtr.Zero;
            }
            // Next try file because that could save us some P/Invokes.
            try
            {
                if (InitWithFile(path, StorageMode.Read))
                    return true;
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }
            try
            {
                pobj = ItemStreamHelper.IShellItemFromPath(path);
                if (pobj != IntPtr.Zero
                    && InitWithItem(pobj, StorageMode.Read))
                    return true;
                if (exceptions.Count == 0)
                    throw new NotSupportedException("The object cannot be initialized at all.");
            }
            catch (Exception ex)
            {
                exceptions.Add(ex);
            }
            finally
            {
                if (pobj != IntPtr.Zero)
                    ItemStreamHelper.ReleaseObject(pobj);
                pobj = IntPtr.Zero;
            }
            throw new AggregateException(exceptions);
        }

        /// <summary>
        /// Calls IPreviewHandler.SetWindow.
        /// </summary>
        public bool ResetWindow()
        {
            EnsureNotDisposed();
            EnsureInitialized();
            var hr = previewHandler.SetWindow(site.WindowHandle, site.PreviewerBounds);
            return (int)hr >= 0;
        }

        /// <summary>
        /// Calls IPreviewHandler.SetRect.
        /// </summary>
        public bool ResetBounds()
        {
            EnsureNotDisposed();
            EnsureInitialized();
            var hr = previewHandler.SetRect(site.PreviewerBounds);
            return (int)hr >= 0;
        }

        /// <summary>
        /// Sets the background if the handler implements IPreviewHandlerVisuals.
        /// </summary>
        /// <param name="color">The background color.</param>
        /// <returns>Whether the call succeeds.</returns>
        public bool SetBackground(Color color)
        {
            var hr = visuals?.SetBackgroundColor(ColorRefFromColor(color));
            return hr.HasValue && (int)hr.Value >= 0;
        }

        /// <summary>
        /// Sets the text color if the handler implements IPreviewHandlerVisuals.
        /// </summary>
        /// <param name="color">The text color.</param>
        /// <returns>Whether the call succeeds.</returns>
        public bool SetForeground(Color color)
        {
            var hr = visuals?.SetTextColor(ColorRefFromColor(color));
            return hr.HasValue && (int)hr.Value >= 0;
        }

        /// <summary>
        /// Sets the font if the handler implements IPreviewHandlerVisuals.
        /// </summary>
        /// <param name="font">The LogFontW reference.</param>
        /// <returns>Whether the call succeeds.</returns>
        public bool SetFont(ref LogFontW font)
        {
            var hr = visuals?.SetFont(ref font);
            return hr.HasValue && (int)hr.Value >= 0;
        }

        /// <summary>
        /// Shows the preview if the object has been successfully initialized.
        /// </summary>
        public void DoPreview()
        {
            EnsureNotDisposed();
            EnsureInitialized();
            EnsureNotShown();
            ResetWindow();
            previewHandler.DoPreview();
            shown = true;
        }

        /// <summary>
        /// Tells the preview handler to set focus to itself.
        /// </summary>
        public void Focus()
        {
            EnsureNotDisposed();
            EnsureInitialized();
            EnsureShown();
            previewHandler.SetFocus();
        }

        /// <summary>
        /// Tells the preview handler to query focus.
        /// </summary>
        /// <returns>The focused window.</returns>
        public IntPtr QueryFocus()
        {
            EnsureNotDisposed();
            EnsureInitialized();
            EnsureShown();
            IntPtr result;
            var hr = previewHandler.QueryFocus(out result);
            if ((int)hr < 0)
                return IntPtr.Zero;
            return result;
        }

        /// <summary>
        /// Unloads the preview and disposes the object. This method is idempotent.
        /// </summary>
        public void UnloadPreview()
        {
            Dispose(true);
        }

        void EnsureNotDisposed()
        {
            if (disposed)
                throw new ObjectDisposedException("PreviewHandler");
        }

        void EnsureInitialized()
        {
            if (!init)
                throw new InvalidOperationException("Object must be initialized before calling this method.");
        }

        void EnsureNotInitialized()
        {
            if (init)
                throw new InvalidOperationException("Object is already initialized and cannot be initialized again.");
        }

        void EnsureShown()
        {
            if (!shown)
                throw new InvalidOperationException("The preview handler must be shown to call this method.");
        }

        void EnsureNotShown()
        {
            if (shown)
                throw new InvalidOperationException("The preview handler must not be shown to call this method.");
        }

        #region IDisposable pattern

        void Dispose(bool disposing)
        {
            if (disposed)
                return;
            disposed = true;
            init = false;
            if (disposing)
            {
                previewHandler.Unload();
                comSite.Dispose();
                site = null;
                Marshal.ReleaseComObject(previewHandler);
            }
            else
            {
                // Field previewHandler must not be used when called from the finalizer.
                var ph = (IPreviewHandler)Marshal.GetUniqueObjectForIUnknown(pPreviewHandler);
                ph.Unload();
                Marshal.ReleaseComObject(previewHandler);
            }
            Marshal.Release(pPreviewHandler);
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
