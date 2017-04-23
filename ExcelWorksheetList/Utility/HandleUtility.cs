using System;
using System.Collections.Generic;

namespace Toybox.Utility
{
	public sealed class HandleUtility
	{

		//private static readonly string CLASS_NAME_XLMAIN = "XLMAIN";
		private static readonly string CLASS_NAME_XLDESK = "XLDESK";
		private static readonly string CLASS_NAME_EXCEL2 = "EXCEL2";
		private static readonly string CLASS_NAME_EXCEL7 = "EXCEL7";
		private static readonly string CLASS_NAME_BAR = "MsoCommandBar";

		public static IntPtr GetMsoCommandBar(IntPtr hWnd)
		{
			if (hWnd == IntPtr.Zero)
			{
				throw new ArgumentException("hWnd");
			}

			return User32Utility.FindWindowEx(hWnd, IntPtr.Zero, CLASS_NAME_BAR, null);
		}

		public static IntPtr GetXLDesk(IntPtr hWnd)
		{
			if (hWnd == IntPtr.Zero)
			{
				throw new ArgumentException("hWnd");
			}

			return User32Utility.FindWindowEx(hWnd, IntPtr.Zero, CLASS_NAME_XLDESK, null);
		}

		public static List<IntPtr> GetExcel2(IntPtr hWnd)
		{
			if (hWnd == IntPtr.Zero)
			{
				throw new ArgumentException("hWnd");
			}

			var list = new List<IntPtr>();
			var prevHWnd = IntPtr.Zero;

			for (;;)
			{
				prevHWnd = User32Utility.FindWindowEx(hWnd, prevHWnd, CLASS_NAME_EXCEL2, null);
				if (prevHWnd == IntPtr.Zero)
				{
					break;
				}
				else
				{
					list.Add(prevHWnd);
				}
			}

			return list;
		}

		public static List<IntPtr> GetExcel7(IntPtr hWnd)
		{
			if (hWnd == IntPtr.Zero)
			{
				throw new ArgumentException("hWnd");
			}

			var list = new List<IntPtr>();
			var prevHWnd = IntPtr.Zero;

			for (;;)
			{
				prevHWnd = User32Utility.FindWindowEx(hWnd, prevHWnd, CLASS_NAME_EXCEL7, null);
				if(prevHWnd == IntPtr.Zero)
				{
					break;
				}
				else
				{
					list.Add(prevHWnd);
				}
			}

			return list;
		}

	}
}
