using System;
using System.Diagnostics;
using Office = Microsoft.Office.Core;

namespace ExcelWorksheetList_2013_2016
{
	public partial class ThisAddIn
	{

		#region VSTO で生成されたコード

		/// <summary>
		/// デザイナーのサポートに必要なメソッドです。
		/// このメソッドの内容をコード エディターで変更しないでください。
		/// </summary>
		private void InternalStartup()
		{
			this.Startup += new EventHandler(ThisAddIn_Startup);
			this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
		}

		#endregion


		#region Protected Methods

		protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			return AppManager.Instance.Ribbon;
		}

		#endregion Protected Methods

		#region Private Methods

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			Debug.WriteLine("ThisAddIn_Startup");

			AppManager.Instance.Startup(this.Application, this.CustomTaskPanes);
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			AppManager.Instance.Shutdown();
		}

		#endregion Private Methods

	}
}
