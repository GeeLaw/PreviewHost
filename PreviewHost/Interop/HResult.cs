namespace PreviewHost.Interop
{
    public enum HResult : int
    {
        S_OK = 0,
        S_FALSE = 1,
        E_ABORT = -2147467260,
        E_ACCESSDENIED = -2147024891,
        E_FAIL = -2147467259,
        E_HANDLE = -2147024890,
        E_INVALIDARG = -2147024809,
        E_NOINTERFACE = -2147467262,
        E_NOTIMPL = -2147467263,
        E_OUTOFMEMORY = -2147024882,
        E_POINTER = -2147467261,
        E_UNEXPECTED = -2147418113,
        CO_E_SERVER_EXEC_FAILURE = -2146959355,
        REGDB_E_CLASSNOTREG = -2147221164
    }
}
