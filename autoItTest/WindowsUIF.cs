using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Drawing.Imaging;
using System.Diagnostics;
using System.ComponentModel;
using System.Data;
using Microsoft.Win32;

namespace ExPlus
{


    public class GetTextTestClass
    {
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);


        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, nuint wParam, StringBuilder lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, nuint wParam, [MarshalAs(UnmanagedType.LPWStr)] string lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, nuint wParam, ref nint lParam);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        static extern IntPtr SendMessage(IntPtr hWnd, uint Msg, nuint wParam, nint lParam);

        const int WM_GETTEXT = 0x000D;
        const int WM_GETTEXTLENGTH = 0x000E;
        const int WM_SETTEXT = 0x000C;

  

    }

    public class WindowHandleInfo
    {
        private delegate bool EnumWindowProc(IntPtr hwnd, IntPtr lParam);

        [DllImport("user32")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool EnumChildWindows(IntPtr window, EnumWindowProc callback, IntPtr lParam);

        private IntPtr _MainHandle;

        WindowHandleInfo(IntPtr handle)
        {
            this._MainHandle = handle;
        }

        List<IntPtr> GetAllChildHandles()
        {
            List<IntPtr> childHandles = new List<IntPtr>();

            GCHandle gcChildhandlesList = GCHandle.Alloc(childHandles);
            IntPtr pointerChildHandlesList = GCHandle.ToIntPtr(gcChildhandlesList);

            try
            {
                EnumWindowProc childProc = new EnumWindowProc(EnumWindow);
                EnumChildWindows(this._MainHandle, childProc, pointerChildHandlesList);
            }
            finally
            {
                gcChildhandlesList.Free();
            }

            return childHandles;
        }

        bool EnumWindow(IntPtr hWnd, IntPtr lParam)
        {
            GCHandle gcChildhandlesList = GCHandle.FromIntPtr(lParam);

            if (gcChildhandlesList == null || gcChildhandlesList.Target == null)
            {
                return false;
            }

            List<IntPtr> childHandles = gcChildhandlesList.Target as List<IntPtr>;
            childHandles.Add(hWnd);

            return true;
        }



        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);


        internal delegate int WindowEnumProc(IntPtr hwnd, IntPtr lparam);

        [DllImport("user32.dll")]
        internal static extern bool EnumChildWindows(IntPtr hwnd, WindowEnumProc func, IntPtr lParam);

     

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);


        static bool doit(IntPtr hwnd, IntPtr lParam)
        {
            Console.WriteLine(lParam);

            return true;
        }

        static readonly string SCN_treeview  = "SysTreeView32";
        static readonly string SCN_listview  = "SysListView32";
        static readonly string SCN_header    = "SysHeader32";
        static readonly string SCN_button    = "Button";
        static readonly string SCN_Edit      = "Edit";
        static readonly string SCN_type1     = "Afx:0000000140000000:8b:0000000000010003:0000000000000000:0000000000000000";
        static readonly string SCN_type2     = "Afx:0000000140000000:83:0000000000010003:0000000000000000:0000000000000000";
        static readonly string SCN_type3     = "Afx:0000000140000000:8:0000000000010003:0000000000000000:0000000000000000";


        public static IntPtr GetEXHandle_treeview   (IntPtr hw, int instance) { return getHandle(hw, SCN_treeview, instance); }
        public static IntPtr GetEXHandle_listview   (IntPtr hw, int instance) { return getHandle(hw, SCN_listview, instance); }
        public static IntPtr GetEXHandle_header     (IntPtr hw, int instance) { return getHandle(hw, SCN_header, instance); }
        public static IntPtr GetEXHandle_button     (IntPtr hw, int instance) { return getHandle(hw, SCN_button, instance); }
        public static IntPtr GetEXHandle_Edit       (IntPtr hw, int instance) { return getHandle(hw, SCN_Edit, instance); }
        public static IntPtr GetEXHandle_type1      (IntPtr hw, int instance) { return getHandle(hw, SCN_type1, instance); }
        public static IntPtr GetEXHandle_type2      (IntPtr hw, int instance) { return getHandle(hw, SCN_type2, instance); }
        public static IntPtr GetEXHandle_type3      (IntPtr hw, int instance) { return getHandle(hw, SCN_type3, instance); }


        static IntPtr getHandle(IntPtr hw, string cn)
        {
            var chhs = getChildHandles(hw);
            var sb = new StringBuilder(1000);
            foreach (var child in chhs)
            {
                GetClassName(child, sb, 1000);
                if (cn.Equals(sb.ToString()))
                    return child;
            }
            return new IntPtr(0);
        }

        static IntPtr getHandle(IntPtr hw, string cn, int instance)
        {
            var chhs = getChildHandles(hw);
            var sb = new StringBuilder(1000);
            int cnt = 0;
            foreach (var child in chhs)
            {
                GetClassName(child, sb, 1000);
                if (cn.Equals(sb.ToString()))
                {
                    cnt++;
                    if (cnt == instance)
                    {
                        return child;
                    }
                }
            }
            return new IntPtr(0);
        }

        static List<IntPtr> getChildHandles(IntPtr hw)
        {
            Process[] localAll = Process.GetProcesses();
            Process app = null;

            foreach (var p in localAll)
            {
                if (p.MainWindowHandle == hw)
                {
                    app = p;
                }
            }
            if (app != null)
            {
                
                return new WindowHandleInfo(app.MainWindowHandle)
                    .GetAllChildHandles();
            }
            return new List<IntPtr>();
        }
    }
    public class AutoItM
    {

        internal delegate int WindowEnumProc(IntPtr hwnd, IntPtr lparam);

        [DllImport("user32.dll")]
        internal static extern bool EnumChildWindows(IntPtr hwnd, WindowEnumProc func, IntPtr lParam);




        [DllImport("user32.dll")]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        public static IntPtr GetHandleWindow(string title)
        {
            return FindWindow(null, title);
        }
        //.net
        public static IntPtr GetWindowHandleByTitle(string wName)
        {
            IntPtr hWnd = IntPtr.Zero;
            foreach (Process pList in Process.GetProcesses())
            {
                Console.WriteLine(pList.ProcessName);
                if (pList.MainWindowTitle.Contains(wName, StringComparison.InvariantCultureIgnoreCase))
                {
                    hWnd = pList.MainWindowHandle;
                }
            }
            return hWnd;
        }
    }

    public class MouseOperations
    {
        [Flags]
        public enum MouseEventFlags
        {
            LeftDown = 0x00000002,
            LeftUp = 0x00000004,
            MiddleDown = 0x00000020,
            MiddleUp = 0x00000040,
            Move = 0x00000001,
            Absolute = 0x00008000,
            RightDown = 0x00000008,
            RightUp = 0x00000010
        }

        [DllImport("user32.dll", EntryPoint = "SetCursorPos")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool SetCursorPos(int x, int y);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetCursorPos(out MousePoint lpMousePoint);

        [DllImport("user32.dll")]
        private static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        public static void SetCursorPosition(int x, int y)
        {
            SetCursorPos(x, y);
        }

        public static void SetCursorPosition(MousePoint point)
        {
            SetCursorPos(point.X, point.Y);
        }

        public static MousePoint GetCursorPosition()
        {
            MousePoint currentMousePoint;
            var gotPoint = GetCursorPos(out currentMousePoint);
            if (!gotPoint) { currentMousePoint = new MousePoint(0, 0); }
            return currentMousePoint;
        }

        public static void MouseEvent(MouseEventFlags value)
        {
            MousePoint position = GetCursorPosition();

            mouse_event
                ((int)value,
                 position.X,
                 position.Y,
                 0,
                 0)
                ;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MousePoint
        {
            public int X;
            public int Y;

            public MousePoint(int x, int y)
            {
                X = x;
                Y = y;
            }
        }
    }


    internal class WindowsUIF
    {

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            private int _Left;
            private int _Top;
            private int _Right;
            private int _Bottom;

            public RECT(RECT Rectangle) : this(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom)
            {
            }
            public RECT(int Left, int Top, int Right, int Bottom)
            {
                _Left = Left;
                _Top = Top;
                _Right = Right;
                _Bottom = Bottom;
            }

            public int X
            {
                get { return _Left; }
                set { _Left = value; }
            }
            public int Y
            {
                get { return _Top; }
                set { _Top = value; }
            }
            public int Left
            {
                get { return _Left; }
                set { _Left = value; }
            }
            public int Top
            {
                get { return _Top; }
                set { _Top = value; }
            }
            public int Right
            {
                get { return _Right; }
                set { _Right = value; }
            }
            public int Bottom
            {
                get { return _Bottom; }
                set { _Bottom = value; }
            }
            public int Height
            {
                get { return _Bottom - _Top; }
                set { _Bottom = value + _Top; }
            }
            public int Width
            {
                get { return _Right - _Left; }
                set { _Right = value + _Left; }
            }
            public Point Location
            {
                get { return new Point(Left, Top); }
                set
                {
                    _Left = value.X;
                    _Top = value.Y;
                }
            }
            public Size Size
            {
                get { return new Size(Width, Height); }
                set
                {
                    _Right = value.Width + _Left;
                    _Bottom = value.Height + _Top;
                }
            }

            public static implicit operator Rectangle(RECT Rectangle)
            {
                return new Rectangle(Rectangle.Left, Rectangle.Top, Rectangle.Width, Rectangle.Height);
            }
            public static implicit operator RECT(Rectangle Rectangle)
            {
                return new RECT(Rectangle.Left, Rectangle.Top, Rectangle.Right, Rectangle.Bottom);
            }
            public static bool operator ==(RECT Rectangle1, RECT Rectangle2)
            {
                return Rectangle1.Equals(Rectangle2);
            }
            public static bool operator !=(RECT Rectangle1, RECT Rectangle2)
            {
                return !Rectangle1.Equals(Rectangle2);
            }

            public override string ToString()
            {
                return "{Left: " + _Left + "; " + "Top: " + _Top + "; Right: " + _Right + "; Bottom: " + _Bottom + "}";
            }

            public override int GetHashCode()
            {
                return ToString().GetHashCode();
            }

            public bool Equals(RECT Rectangle)
            {
                return Rectangle.Left == _Left && Rectangle.Top == _Top && Rectangle.Right == _Right && Rectangle.Bottom == _Bottom;
            }

            public override bool Equals(object Object)
            {
                if (Object is RECT)
                {
                    return Equals((RECT)Object);
                }
                else if (Object is Rectangle)
                {
                    return Equals(new RECT((Rectangle)Object));
                }

                return false;
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        static public extern bool MoveWindow(IntPtr hWnd, int X, int Y, int nWidth, int nHeight, bool bRepaint);

        [DllImport("user32.dll")]
        public static extern IntPtr GetDC(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern Int32 ReleaseDC(IntPtr hwnd, IntPtr hdc);

        [DllImport("gdi32.dll")]
        static extern uint GetPixel(IntPtr hdc, int nXPos, int nYPos);

        static public Color GetPixelColor(IntPtr hwnd, int x, int y)
        {
            IntPtr hdc = GetDC(hwnd);
            uint pixel = GetPixel(hdc, x, y);
            //Console.WriteLine(pixel);
            ReleaseDC(hwnd, hdc);
            Color color = Color.FromArgb((int)(pixel & 0x000000FF),
                            (int)(pixel & 0x0000FF00) >> 8,
                            (int)(pixel & 0x00FF0000) >> 16);
            return color;
        }




        [DllImport("user32.dll")]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);
        [DllImport("user32.dll")]
        public static extern bool PrintWindow(IntPtr hWnd, IntPtr hdcBlt, int nFlags);

        public static Bitmap PrintWindow(IntPtr hwnd)
        {
            RECT rc;
            GetWindowRect(hwnd, out rc);



            Bitmap bmp = new(rc.Width, rc.Height, PixelFormat.Format32bppArgb);
            Graphics gfxBmp = Graphics.FromImage(bmp);
            IntPtr hdcBitmap = gfxBmp.GetHdc();

            PrintWindow(hwnd, hdcBitmap, 0);

            gfxBmp.ReleaseHdc(hdcBitmap);
            gfxBmp.Dispose();

            return bmp;
        }
        public static int getPixelEvenOut(IntPtr hwnd, int x, int y)
        {
            var b = PrintWindow(hwnd);
            return b.GetPixel(x, y).ToArgb();
        }


        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);


        private struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public int showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

    }
}
