using System.Runtime.InteropServices;
using System.Text;
using System.Drawing;

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

        [DllImport("shell32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr SHGetFileInfo(IntPtr pszPath, uint dwFileAttributes, ref SHFILEINFO psfi, uint cbSizeFileInfo, uint uFlags);

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
        internal static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("imm32.dll")]
        internal static extern IntPtr ImmGetContext(IntPtr hWnd);

        [DllImport("imm32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ImmGetOpenStatus(IntPtr hIMC);

        [DllImport("imm32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ImmGetConversionStatus(IntPtr hIMC, out int lpfdwConversion, out int lpfdwSentence);

        [DllImport("imm32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ImmSetOpenStatus(IntPtr hIMC, bool fOpen);

        [DllImport("imm32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ImmSetConversionStatus(IntPtr hIMC, int fdwConversion, int fdwSentence);

        [DllImport("imm32.dll", CharSet = CharSet.Unicode)]
        internal static extern uint ImmGetImeMenuItemsW(IntPtr hIMC, uint dwFlags, uint dwType, ref IMEMENUITEMINFOW lpImeParentMenu, [Out] IMEMENUITEMINFOW[]? lpImeMenu, uint dwSize);

        [DllImport("imm32.dll", CharSet = CharSet.Unicode)]
        internal static extern uint ImmGetImeMenuItemsW(IntPtr hIMC, uint dwFlags, uint dwType, IntPtr lpImeParentMenu, [Out] IMEMENUITEMINFOW[]? lpImeMenu, uint dwSize);

        [DllImport("imm32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ImmNotifyIME(IntPtr hIMC, uint dwAction, uint dwIndex, uint dwValue);

        [DllImport("imm32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool ImmReleaseContext(IntPtr hWnd, IntPtr hIMC);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetShellWindow();

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        internal static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

        [DllImport("user32.dll")]
        internal static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

        internal const uint GW_OWNER = 4;

        [DllImport("user32.dll")]
        internal static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        internal static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern short GetAsyncKeyState(int vKey);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr SetWindowsHookEx(int idHook, LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        internal static extern IntPtr CreatePopupMenu();

        [DllImport("user32.dll")]
        internal static extern bool DestroyMenu(IntPtr hMenu);

        // AppBar API
        [DllImport("shell32.dll")]
        internal static extern IntPtr SHAppBarMessage(uint dwMessage, ref APPBARDATA pData);

        [StructLayout(LayoutKind.Sequential)]
        internal struct APPBARDATA
        {
            public uint cbSize;
            public IntPtr hWnd;
            public uint uCallbackMessage;
            public uint uEdge;
            public RECT rc;
            public IntPtr lParam;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;

            public RECT(int left, int top, int right, int bottom)
            {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public Rectangle ToRectangle() => Rectangle.FromLTRB(Left, Top, Right, Bottom);
        }

        internal enum AppBarMsg
        {
            New = 0x00000001,
            Remove = 0x00000002,
            SetPos = 0x00000003,
            GetPos = 0x00000004,
            GetAutoHideBar = 0x00000005,
            SetAutoHideBar = 0x00000006,
            WindowPosChanged = 0x00000009,
            SetState = 0x0000000A
        }

        internal enum AppBarEdge
        {
            Left = 0,
            Top = 1,
            Right = 2,
            Bottom = 3
        }

        internal enum AppBarState
        {
            AutoHide = 0x00000001,
            AlwaysOnTop = 0x00000002,
            DockTop = 0x00000004,
            DockBottom = 0x00000008,
            DockLeft = 0x00000010,
            DockRight = 0x00000020
        }

        // AppBar notification codes (wParam values for callback message)
        internal const int ABN_STATECHANGE = 0x00000000;
        internal const int ABN_POSCHANGED = 0x00000001;
        internal const int ABN_FULLSCREENAPPNOTIFY = 0x00000002;
        internal const int ABN_WINDOWARRANGE = 0x00000003;

        // Constants for context menu
        internal const uint TPM_RETURNCMD = 0x0100;
        internal const uint TPM_LEFTBUTTON = 0x0000;
        internal const int WH_KEYBOARD_LL = 13;
        internal const int WH_MOUSE_LL = 14;
        internal const int WM_KEYDOWN = 0x0100;
        internal const int WM_KEYUP = 0x0101;
        internal const int WM_SYSKEYDOWN = 0x0104;
        internal const int WM_SYSKEYUP = 0x0105;
        internal const int WM_MOUSEACTIVATE = 0x0021;
        internal const int WM_ACTIVATEAPP = 0x001C;
        internal const int WM_USER = 0x0400;
        internal const int WM_LBUTTONDOWN = 0x0201;
        internal const int WM_LBUTTONUP = 0x0202;
        internal const int WM_RBUTTONDOWN = 0x0204;
        internal const int WM_RBUTTONUP = 0x0205;
        internal const int WM_MBUTTONDOWN = 0x0207;
        internal const int WM_MBUTTONUP = 0x0208;
        internal const int IME_CMODE_NATIVE = 0x0001;
        internal const int IME_CMODE_KATAKANA = 0x0002;
        internal const int IME_CMODE_FULLSHAPE = 0x0008;
        internal const int IME_CMODE_ROMAN = 0x0010;
        internal const uint NI_IMEMENUSELECTED = 0x0018;
        internal const uint IGIMIF_RIGHTMENU = 0x0001;
        internal const uint IGIMII_CMODE = 0x0001;
        internal const uint IGIMII_SMODE = 0x0002;
        internal const uint IGIMII_CONFIGURE = 0x0004;
        internal const uint IGIMII_TOOLS = 0x0008;
        internal const uint IGIMII_HELP = 0x0010;
        internal const uint IGIMII_OTHER = 0x0020;
        internal const uint IGIMII_INPUTTOOLS = 0x0040;
        internal const uint IMFT_RADIOCHECK = 0x0001;
        internal const uint IMFT_SEPARATOR = 0x0002;
        internal const uint IMFT_SUBMENU = 0x0004;
        internal const uint IMFS_GRAYED = 0x0003;
        internal const uint IMFS_DISABLED = 0x0003;
        internal const uint IMFS_CHECKED = 0x0008;
        internal const uint IMFS_DEFAULT = 0x1000;
        internal const int VK_MENU = 0x12;
        internal const int VK_LWIN = 0x5B;
        internal const int VK_RWIN = 0x5C;
        internal const int SW_RESTORE = 9;
        internal const int SW_MINIMIZE = 6;

        internal static int LOWORD(int value) => value & 0xFFFF;
        internal static int HIWORD(int value) => (value >> 16) & 0xFFFF;
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

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        internal struct IMEMENUITEMINFOW
        {
            public uint cbSize;
            public uint fType;
            public uint fState;
            public uint wID;
            public IntPtr hbmpChecked;
            public IntPtr hbmpUnchecked;
            public uint dwItemData;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string szString;
            public IntPtr hbmpItem;
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

        internal delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);

        internal delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);
        internal delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        internal struct KBDLLHOOKSTRUCT
        {
            public uint vkCode;
            public uint scanCode;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct MSLLHOOKSTRUCT
        {
            public Point pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }
    }
}
