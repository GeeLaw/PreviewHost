using System;
using System.Collections.Generic;
using System.Windows.Interop;

namespace PreviewHost.Interop
{
    /// <summary>
    /// Represents a managed preview handler frame. GetAcceleratorTable and TranslateAccelerator might be called from a different thread.
    /// </summary>
    public interface IPreviewHandlerManagedFrame
    {
        /// <summary>
        /// Enumerates the accelerator table. The method must not throw any exception other than OutOfMemoryException. The result can contain at most 32767 items.
        /// </summary>
        /// <returns>The enumerable of ACCEL.</returns>
        IEnumerable<AcceleratorEntry> GetAcceleratorTable();
        /// <summary>
        /// Tries to process an accelerator. This method must not throw any exception other than OutOfMemoryException.
        /// </summary>
        /// <param name="msg">The message. Do not modify this message. It is passed as reference only for performance reason.</param>
        /// <returns>If the accelerator is processed, true; otherwise, false.</returns>
        bool TranslateAccelerator(ref MSG msg);
        /// <summary>
        /// Gets the window handle.
        /// </summary>
        IntPtr WindowHandle { get; }
        /// <summary>
        /// Gets the previewer bounds.
        /// </summary>
        Rect PreviewerBounds { get; }
    }
}
