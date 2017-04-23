using Microsoft.Office.Tools;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Toybox.Core;
using Toybox.ExcelWorksheetList.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorksheetList_2010
{
	public class AppManager
	{

		#region [Static]

		public static AppManager Instance
		{
			get { return _instance ?? (_instance = new AppManager()); }
		}
		private static AppManager _instance;

		private static readonly string TITLE = "Sheet List";

		#endregion [Static]


		#region Constructor

		private AppManager()
		{
			this.CM = new ControlManager();
			this.VMM = new ViewModelManager();
		}

		#endregion Constructor


		#region Public Members

		public Office.IRibbonExtensibility Ribbon { get { return this.CM.Ribbon; } }

		#endregion Public Members

		#region Protected Members

		protected Excel.Application App { get; set; }
		protected CustomTaskPane Pane { get; set; }

		protected ControlManager CM { get; set; }
		protected ViewModelManager VMM { get; set; }

		protected System.Threading.Timer Timer { get; set; }

		#endregion Protected Members


		#region Public Methods

		public void Startup(Excel.Application app, CustomTaskPaneCollection panes)
		{
			Debug.WriteLine("Startup");

			this.App = app;
			this.Pane = panes.Add(CM.ContainerControl, TITLE);
			this.VMM.SheetListControl.Workbook = this.App.ActiveWorkbook;

			// Hook
			this.Hook(this.VMM.SheetListControl.Workbook);

			// App
			this.App.WorkbookActivate += Application_WorkbookActivate;
			this.App.WorkbookDeactivate += Application_WorkbookDeactivate;

			// ViewModel
			this.VMM.SheetListControl.Workbook = this.VMM.SheetListControl.Workbook;

			// Control
			this.CM.ContainerControl.DataContext = this.VMM.SheetListControl;

			this.CM.Ribbon.VisibilityChanged += Ribbon_VisibilityChanged;
			this.CM.Ribbon.Visible = true;

			// Pane
			this.Pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
			this.Pane.Visible = true;
			this.Pane.VisibleChanged += Pane_VisibleChanged;
		}

		public void Shutdown()
		{
			Debug.WriteLine("Shutdown");

			// Unhook
			this.Unhook(this.VMM.SheetListControl.Workbook);

			// App
			this.App.WorkbookActivate -= Application_WorkbookActivate;
			this.App.WorkbookDeactivate -= Application_WorkbookDeactivate;
			this.Pane.VisibleChanged -= Pane_VisibleChanged;

			// VieModel
			this.VMM.SheetListControl.Workbook = null;

			// Control
			this.CM.ContainerControl.Dispose();
		}

		#endregion Public Methods

		#region Private Methods

		private void Application_WorkbookActivate(Excel.Workbook workbook)
		{
			Debug.WriteLine("Workbook Activate");

			//this.Unhook(this.VMM.SheetListControl.Workbook);
			this.VMM.SheetListControl.Workbook = workbook;
			this.Hook(this.VMM.SheetListControl.Workbook);
		}

		private void Application_WorkbookDeactivate(Excel.Workbook workbook)
		{
			Debug.WriteLine("Workbook Deactivate");

			this.VMM.SheetListControl.Workbook = null;
			this.Unhook(workbook);
		}

		private void Ribbon_VisibilityChanged(object sender, EventArgs<bool> e)
		{
			this.Pane.Visible = e.Item;
		}

		private void Pane_VisibleChanged(object sender, EventArgs e)
		{
			this.CM.Ribbon.Visible = this.Pane.Visible;
		}

		private void Hook(Excel.Workbook workbook)
		{
			if (workbook == null)
			{
				return;
			}

			workbook.SheetActivate += this.Workbook_SheetActivate;
			workbook.NewSheet += this.Workbook_NewSheet;
			workbook.NewChart += this.Workbook_NewSheet;
			workbook.AfterSave += this.Workbook_AfterSave;

			//this.CM.XLMain = new XLMainControl((IntPtr)workbook.Application.Hwnd);
			//this.CM.XLMain.Changed += this.XLMain_Changed;

			this.VMM.SheetListControl?.Update();
		}

		private void Unhook(Excel.Workbook workbook)
		{
			if (workbook != null)
			{
				workbook.SheetActivate -= this.Workbook_SheetActivate;
			}

			//if (this.CM.XLMain != null)
			//{
			//	this.CM.XLMain.Changed -= this.XLMain_Changed;
			//	this.CM.XLMain.Dispose();
			//	this.CM.XLMain = null;
			//}

			this.VMM.SheetListControl?.Update();

			this.Timer?.Dispose();
		}

		private void Workbook_SheetActivate(object sheet)
		{
			Debug.WriteLine("Sheet Activate");

			this.VMM.SheetListControl?.WorkbookInfo?.Select(sheet);
		}

		private void Workbook_NewSheet(object sheet)
		{
			this.VMM.SheetListControl?.Update();
		}

		private void Workbook_AfterSave(bool Success)
		{
			this.VMM.SheetListControl?.Update();
		}

		private void XLMain_Changed(object sender, EventArgs e)
		{
			Debug.WriteLine("Changeing");

			this.Timer?.Dispose();
			this.Timer = new System.Threading.Timer((_) =>
			{
				Debug.WriteLine("Changed");

				this.CM.ContainerControl?.Invoke((MethodInvoker)delegate ()
				{
					this.VMM.SheetListControl?.Update();
				});

			}, null, 50, System.Threading.Timeout.Infinite);
		}

		#endregion Private Methods

	}
}
