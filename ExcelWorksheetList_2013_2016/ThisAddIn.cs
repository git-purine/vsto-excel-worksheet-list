using System;
using Office = Microsoft.Office.Core;
using System.Diagnostics;

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

		#region Private Members

		private AppManager appManager { get; } = new AppManager();

		#endregion Private Members


		#region Protected Methods

		protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			return appManager.Ribbon;
		}

		#endregion Protected Methods

		#region Private Methods

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			this.Application.WindowActivate += Application_WindowActivate;

			this.Application.WorkbookOpen += Application_WorkbookOpen;
			this.Application.WorkbookBeforeClose += Application_WorkbookBeforeClose;

			appManager.Startup(this.Application, this.CustomTaskPanes);
		}

		private void Application_WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook Wb, ref bool Cancel)
		{


		}

		private void Application_WorkbookOpen(Microsoft.Office.Interop.Excel.Workbook Wb)
		{
			Debug.WriteLine(Wb.Windows.Count);
			foreach(Microsoft.Office.Interop.Excel.Window window in Wb.Windows)
			{
				Debug.WriteLine(window.Index);
			}
		}

		private void Application_WindowActivate(Microsoft.Office.Interop.Excel.Workbook Wb, Microsoft.Office.Interop.Excel.Window Wn)
		{
			//var form = new System.Windows.Forms.UserControl();
			//var pane = this.CustomTaskPanes.Add(form, "test", Wn);

			//pane.Visible = true;
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			appManager.Shutdown();
		}

		#endregion Private Methods

	}
}
