Add-Type @'
using System;using System.Text;using System.Runtime.InteropServices;
public static class GifMenuNative{
 public delegate bool Call(IntPtr h,IntPtr x);
 [StructLayout(LayoutKind.Sequential)]public struct R{public int L,T,Rt,B;}
 [StructLayout(LayoutKind.Sequential)]public struct P{public int X,Y;}
 [DllImport("user32.dll")]public static extern IntPtr SetThreadDpiAwarenessContext(IntPtr c);
 [DllImport("user32.dll")]public static extern bool EnumWindows(Call c,IntPtr x);
 [DllImport("user32.dll")]public static extern bool IsWindowVisible(IntPtr h);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]public static extern int GetClassName(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll",CharSet=CharSet.Unicode)]public static extern int GetWindowText(IntPtr h,StringBuilder b,int n);
 [DllImport("user32.dll")]public static extern bool SetForegroundWindow(IntPtr h);
 [DllImport("user32.dll")]public static extern IntPtr GetForegroundWindow();
 [DllImport("user32.dll")]public static extern bool GetClientRect(IntPtr h,out R r);
 [DllImport("user32.dll")]public static extern bool ClientToScreen(IntPtr h,ref P p);
 [DllImport("user32.dll")]public static extern bool SetCursorPos(int x,int y);
 [DllImport("user32.dll")]public static extern bool GetCursorPos(out P p);
 [DllImport("user32.dll")]public static extern IntPtr WindowFromPoint(P p);
 [DllImport("user32.dll")]public static extern IntPtr GetAncestor(IntPtr h,uint f);
 [DllImport("user32.dll")]public static extern void mouse_event(uint f,int x,int y,uint d,UIntPtr e);
 [DllImport("user32.dll")]public static extern void keybd_event(byte vk,byte scan,uint flags,UIntPtr e);
 [DllImport("dwmapi.dll")]public static extern int DwmGetWindowAttribute(IntPtr h,int a,out R r,int size);
 public static IntPtr Find(string name){IntPtr found=IntPtr.Zero;EnumWindows((h,x)=>{var c=new StringBuilder(256);var t=new StringBuilder(1024);GetClassName(h,c,256);GetWindowText(h,t,1024);if(IsWindowVisible(h)&&c.ToString().ToLower().Contains("screenclass")&&t.ToString().IndexOf(name,StringComparison.OrdinalIgnoreCase)>=0)found=h;return true;},IntPtr.Zero);return found;}
}
'@
