using System;
using System.Runtime.InteropServices;

namespace PreviewHost.Interop
{
    [StructLayout(LayoutKind.Sequential)]
    public struct AcceleratorEntry
    {
        public byte IsVirtual;
        public ushort Key;
        public ushort Command;

        [DllImport("user32.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        public static extern IntPtr CreateAcceleratorTable([In, MarshalAs(UnmanagedType.LPArray)] AcceleratorEntry[] paccel, int cAccel);
    }
}
