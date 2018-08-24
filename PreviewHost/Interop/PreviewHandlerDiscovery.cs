using System;
using System.Runtime.InteropServices;
using System.Text;

namespace PreviewHost.Interop
{
    public static class PreviewHandlerDiscovery
    {
        [Flags]
        enum AssocF : uint
        {
            InitDefaultToStar = 0x00000004,
            NoTruncate = 0x00000020
        }

        enum AssocStr
        {
            ShellExtension = 16
        }

        [ComImport, Guid("c46ca590-3c3f-11d2-bee6-0000f805ca57"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        interface IQueryAssociations
        {
            [PreserveSig]
            HResult Init([In] AssocF flags, [In, Optional] string pszAssoc, [In, Optional] IntPtr hkProgid, [In, Optional] IntPtr hwnd);
            [PreserveSig]
            HResult GetString([In] AssocF flags, [In] AssocStr str, [In, Optional] string pszExtra, [Out, Optional] StringBuilder pwszOut, [In, Out] ref uint pcchOut);
            // Method slots beyond GetString are ignored since they are not used.
        }

        [DllImport("shlwapi.dll", PreserveSig = true, CallingConvention = CallingConvention.StdCall)]
        static extern HResult AssocCreate(Guid clsid, ref Guid riid, out IntPtr ppv);

        const string IPreviewHandlerIid = "{8895b1c6-b41f-4c1c-a562-0d564250836f}";
        static readonly Guid QueryAssociationsClsid = new Guid(0xa07034fd, 0x6caa, 0x4954, 0xac, 0x3f, 0x97, 0xa2, 0x72, 0x16, 0xf9, 0x8a);
        static readonly Guid IQueryAssociationsIid = Guid.ParseExact("c46ca590-3c3f-11d2-bee6-0000f805ca57", "d");

        public static Guid? FindPreviewHandlerFor(string extension, IntPtr hwnd)
        {
            IntPtr pqa;
            var iid = IQueryAssociationsIid;
            var hr = AssocCreate(QueryAssociationsClsid, ref iid, out pqa);
            if ((int)hr < 0)
                return null;
            var queryAssoc = (IQueryAssociations)Marshal.GetUniqueObjectForIUnknown(pqa);
            Marshal.Release(pqa);
            try
            {
                hr = queryAssoc.Init(AssocF.InitDefaultToStar, extension, IntPtr.Zero, hwnd);
                if ((int)hr < 0)
                    return null;
                var sb = new StringBuilder(128);
                uint cch = 64;
                hr = queryAssoc.GetString(AssocF.NoTruncate, AssocStr.ShellExtension, IPreviewHandlerIid, sb, ref cch);
                if ((int)hr < 0)
                    return null;
                return Guid.Parse(sb.ToString());
            }
            catch
            {
                return null;
            }
            finally
            {
                Marshal.ReleaseComObject(queryAssoc);
            }
        }
    }
}
