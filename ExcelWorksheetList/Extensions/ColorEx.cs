using System.Windows.Media;

namespace Toybox.Extensions
{
	public static class ColorEx
	{

		public static float GetBrightness(this Color color)
		{
			return System.Drawing.Color.FromArgb(color.A, color.R, color.G, color.B).GetBrightness();
		}

	}
}
