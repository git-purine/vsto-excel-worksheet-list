using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using Office = Microsoft.Office.Core;

namespace ExcelWorksheetList_2013_2016.Ribbons
{
	[ComVisible(true)]
	public class WorksheetListRibbon : Office.IRibbonExtensibility
	{
		private Office.IRibbonUI ribbon;

		public WorksheetListRibbon()
		{
		}

		#region IRibbonExtensibility のメンバー

		public string GetCustomUI(string ribbonID)
		{
			return GetResourceText("ExcelWorksheetList_2013_2016.Ribbons.WorksheetListRibbon.xml");
		}

		#endregion

		#region リボンのコールバック

		public void Ribbon_Load(Office.IRibbonUI ribbonUI)
		{
			this.ribbon = ribbonUI;
		}

		public void button_Click(Office.IRibbonControl control)
		{
			Debug.WriteLine("button_Click");

			AppManager.Instance.ShowActiveWorksheetPane();
		}

		#endregion

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
