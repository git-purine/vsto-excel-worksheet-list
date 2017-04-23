using System;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace Toybox.ExcelWorksheetList.Ribbons
{
	using Core;

	[ComVisible(true)]
	public class VisibilityRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		public VisibilityRibbon()
		{
		}

		#region IRibbonExtensibility のメンバー

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("Toybox.ExcelWorksheetList.Ribbons.VisibilityRibbon.xml");
		}

		#endregion

		#region リボンのコールバック

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		public bool Visible_getPressed(Office.IRibbonControl control)
		{
			return this.Visible;
		}

		public void Visible_onAction(Office.IRibbonControl control, bool value)
		{
			this._visible = value;
			this.VisibilityChanged?.Invoke(control, new EventArgs<bool>(value));
		}

		#endregion

		#region Events

		public event EventHandler<EventArgs<bool>> VisibilityChanged;

		#endregion Events

		#region Public Members

		public bool Visible
		{
			get { return this._visible; }
			set
			{
				this._visible = value;
				this.ribbon?.InvalidateControl("ewlVisible");
			}
		}
		private bool _visible;


		#endregion Public Members


		#region ヘルパー

		private static string GetResourceText(string resourceName)
		{
			Assembly asm = Assembly.GetExecutingAssembly();
			string[] resourceNames = asm.GetManifestResourceNames();
			for (int i = 0; i < resourceNames.Length; ++i)
			{
				if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
				{
					using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
					{
						if (resourceReader != null)
						{
							return resourceReader.ReadToEnd();
						}
					}
				}
			}
			return null;
		}

		#endregion
	}
}
