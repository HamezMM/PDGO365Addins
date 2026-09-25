using System;
using System.Windows.Forms;

namespace SpeakerRename.UI
{
    /// <summary>Wraps a raw HWND (e.g. Word's window) so it can own a WinForms dialog.</summary>
    internal sealed class WindowHandle : IWin32Window
    {
        public WindowHandle(IntPtr handle) => Handle = handle;

        public IntPtr Handle { get; }
    }
}
