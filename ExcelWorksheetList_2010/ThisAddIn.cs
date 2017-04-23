﻿using System;
using Office = Microsoft.Office.Core;

namespace ExcelWorksheetList_2010
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
			this.Startup += new System.EventHandler(ThisAddIn_Startup);
			this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
		}

		#endregion

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			AppManager.Instance.Startup(this.Application, this.CustomTaskPanes);
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			AppManager.Instance.Shutdown();
		}

		protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
		{
			return AppManager.Instance.Ribbon;
		}
	}
}
