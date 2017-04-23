using System;
using System.Diagnostics;
using System.Reflection;

namespace Toybox.Extensions
{
	public static class ComObjectEx
	{

		public static object InvokeMember(this object obj, string propertyName, BindingFlags flags = BindingFlags.Public | BindingFlags.GetProperty)
		{
			if (string.IsNullOrWhiteSpace(propertyName))
			{
				throw new ArgumentException(nameof(propertyName));
			}

			try
			{
				return obj.GetType().InvokeMember(propertyName, flags, null, obj, null);
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.Message);
				return null;
			}
		}

		public static object InvokeMethod(this object obj, string methodName, BindingFlags flags = BindingFlags.Public | BindingFlags.InvokeMethod)
		{
			if (string.IsNullOrWhiteSpace(methodName))
			{
				throw new ArgumentException(nameof(methodName));
			}

			try
			{
				return obj.GetType().InvokeMember(methodName, flags, null, obj, null);
			}
			catch (Exception ex)
			{
				Debug.WriteLine(ex.Message);
				return null;
			}
		}

	}
}
