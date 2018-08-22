# PreviewHost

An example [`IPreviewHandlerFrame`](https://docs.microsoft.com/en-us/windows/desktop/api/shobjidl_core/nn-shobjidl_core-ipreviewhandlerframe) implementation using WPF (Windows Presentation Foundation). This project aims to be a correct-to-every-detail implementation of `IPreviewHandlerFrame` in managed code.

Encoding and line terminator choice are kept the default for files created and edited with Visual Studio (2017 Community). Other text files use UTF-8 (without BOM) and LF.

Steps:

1. [Create the UI](https://geelaw.blog/entries/ipreviewhandlerframe-wpf-1-ui-assoc/#create-ui)
