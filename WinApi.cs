using System.Runtime.InteropServices;
using System.Text;

namespace win9xplorer
{
    /// <summary>
    /// Contains Windows API declarations and structures used by the file explorer
    /// </summary>
    internal static class WinApi
    {
        // Windows API declarations for getting file icons
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr SHGetFileInfo(string pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool DestroyIcon(IntPtr hIcon);

        // Additional API for extracting shell32 icons by index
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr ExtractIcon(IntPtr hInst, string lpszExeFileName, int nIconIndex);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern bool GetIconInfo(IntPtr hIcon, out ICONINFO piconinfo);

        [DllImport("gdi32.dll", SetLastError = true)]
        internal static extern bool DeleteObject(IntPtr hObject);

        // Windows API declarations for system context menus
        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr SHGetDesktopFolder(out IntPtr ppshf);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        internal static extern int SHParseDisplayName([MarshalAs(UnmanagedType.LPWStr)] string pszName, IntPtr pbc, out IntPtr ppidl, uint sfgaoIn, out uint psfgaoOut);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        internal static extern int SHBindToParent(IntPtr pidl, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv, out IntPtr ppidlLast);

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        internal static extern int SHBindToObject(IntPtr psf, IntPtr pidl, IntPtr pbc, [MarshalAs(UnmanagedType.LPStruct)] Guid riid, out IntPtr ppv);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern int TrackPopupMenuEx(IntPtr hmenu, uint fuFlags, int x, int y, IntPtr hwnd, IntPtr lptpm);

        [DllImport("ole32.dll")]
        internal static extern void CoTaskMemFree(IntPtr pv);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll")]
        internal static extern bool DestroyMenu(IntPtr hMenu);

        // Constants for context menu
        internal const uint TPM_RETURNCMD = 0x0100;
        internal const uint TPM_LEFTBUTTON = 0x0000;
        internal static readonly Guid IID_IShellFolder = new Guid("000214E6-0000-0000-C000-000000000046");
        internal static readonly Guid IID_IContextMenu = new Guid("000214e4-0000-0000-c000-000000000046");

        // Constants for SHGetFileInfo
        internal const uint SHGFI_ICON = 0x000000100;
        internal const uint SHGFI_DISPLAYNAME = 0x000000200;
        internal const uint SHGFI_TYPENAME = 0x000000400;
        internal const uint SHGFI_ATTRIBUTES = 0x000000800;
        internal const uint SHGFI_ICONLOCATION = 0x000001000;
        internal const uint SHGFI_EXETYPE = 0x000002000;
        internal const uint SHGFI_SYSICONINDEX = 0x000004000;
        internal const uint SHGFI_LINKOVERLAY = 0x000008000;
        internal const uint SHGFI_SELECTED = 0x000010000;
        internal const uint SHGFI_ATTR_SPECIFIED = 0x000020000;
        internal const uint SHGFI_LARGEICON = 0x000000000;
        internal const uint SHGFI_SMALLICON = 0x000000001;
        internal const uint SHGFI_OPENICON = 0x000000002;
        internal const uint SHGFI_SHELLICONSIZE = 0x000000004;
        internal const uint SHGFI_PIDL = 0x000000008;
        internal const uint SHGFI_USEFILEATTRIBUTES = 0x000000010;
        internal const uint FILE_ATTRIBUTE_DIRECTORY = 0x00000010;

        [StructLayout(LayoutKind.Sequential)]
        internal struct ICONINFO
        {
            public bool fIcon;
            public int xHotspot;
            public int yHotspot;
            public IntPtr hbmMask;
            public IntPtr hbmColor;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Auto)]
        internal struct SHFILEINFO
        {
            public IntPtr hIcon;
            public int iIcon;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szDisplayName;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szTypeName;
            public uint dwAttributes;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct CMINVOKECOMMANDINFO
        {
            public int cbSize;
            public uint fMask;
            public IntPtr hwnd;
            public IntPtr lpVerb;
            public IntPtr lpParameters;
            public IntPtr lpDirectory;
            public int nShow;
            public uint dwHotKey;
            public IntPtr hIcon;
        }

        // COM interfaces for shell context menu
        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214E6-0000-0000-C000-000000000046")]
        internal interface IShellFolder
        {
            [PreserveSig]
            int ParseDisplayName(IntPtr hwnd, IntPtr pbc, [MarshalAs(UnmanagedType.LPWStr)] string pszDisplayName, ref uint pchEaten, out IntPtr ppidl, ref uint pdwAttributes);

            [PreserveSig]
            int EnumObjects(IntPtr hwnd, int grfFlags, out IntPtr ppenumIDList);

            [PreserveSig]
            int BindToObject(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv);

            [PreserveSig]
            int BindToStorage(IntPtr pidl, IntPtr pbc, [In] ref Guid riid, out IntPtr ppv);

            [PreserveSig]
            int CompareIDs(IntPtr lParam, IntPtr pidl1, IntPtr pidl2);

            [PreserveSig]
            int CreateViewObject(IntPtr hwndOwner, [In] ref Guid riid, out IntPtr ppv);

            [PreserveSig]
            int GetAttributesOf(uint cidl, [MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl, ref uint rgfInOut);

            [PreserveSig]
            int GetUIObjectOf(IntPtr hwndOwner, uint cidl, [MarshalAs(UnmanagedType.LPArray)] IntPtr[] apidl, [In] ref Guid riid, IntPtr rgfReserved, out IntPtr ppv);

            [PreserveSig]
            int GetDisplayNameOf(IntPtr pidl, int uFlags, IntPtr lpName);

            [PreserveSig]
            int SetNameOf(IntPtr hwndOwner, IntPtr pidl, [MarshalAs(UnmanagedType.LPWStr)] string pszName, int uFlags, out IntPtr ppidlOut);
        }

        [ComImport]
        [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        [Guid("000214e4-0000-0000-c000-000000000046")]
        internal interface IContextMenu
        {
            [PreserveSig]
            int QueryContextMenu(IntPtr hmenu, uint indexMenu, uint idCmdFirst, uint idCmdLast, uint uFlags);

            [PreserveSig]
            int InvokeCommand(IntPtr pici);

            [PreserveSig]
            int GetCommandString(uint idCmd, uint uFlags, IntPtr pwReserved, [MarshalAs(UnmanagedType.LPStr)] StringBuilder pszName, uint cchMax);
        }
    }
}