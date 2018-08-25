using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PreviewHost.Interop
{
    [Flags]
    public enum StorageMode : uint
    {
        Read = 0x0,
        Write = 0x1,
        ReadWrite = 0x2,
        ShareDenyNone = 0x40,
        ShareDenyRead = 0x30,
        ShareDenyWrite = 0x20,
        ShareExclusive = 0x10,
        Priority = 0x40000,
        Create = 0x1000,
        Convert = 0x20000,
        FailIfThere = 0x0,
        Direct = 0x0,
        Transacted = 0x10000,
        NoScratch = 0x100000,
        NoSnapshot = 0x200000,
        Simple = 0x8000000,
        DirectSingleWriterMultipleReaders = 0x400000,
        DeleteOnRelease = 0x4000000
    }
}
