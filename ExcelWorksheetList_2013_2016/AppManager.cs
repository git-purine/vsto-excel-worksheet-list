using Microsoft.Office.Tools;
using System;
using System.Diagnostics;
using System.Windows.Forms;
using Toybox.Core;
using Toybox.ExcelWorksheetList.Controls;
using Toybox.Utility;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorksheetList_2013_2016
{
	public class AppManager
	{

		#region [Static]

		private static readonly string TITLE = "Sheet List";

		#endregion [Static]


		#region Constructor

		public AppManager()
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
			this.App = app;
			this.Pane = panes.Add(CM.ContainerControl, TITLE);

			this.VMM.SheetListControl.Workbook = this.App.ActiveWorkbook;

			// Hook
			this.Hook(this.VMM.SheetListControl.Workbook);

			// App
			this.App.WorkbookActivate += Application_WorkbookActivate;
			this.App.WorkbookDeactivate += Application_WorkbookDeactivate;

			this.App.WindowActivate += App_WindowActivate;
			this.App.WindowDeactivate += App_WindowDeactivate;

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

		private void App_WindowActivate(Excel.Workbook Wb, Excel.Window Wn)
		{
			Debug.WriteLine("Window Activate");
			Debug.WriteLine(Wb.Name);


			Wb.BeforeClose += Wb_BeforeClose;
		}

		private void Wb_BeforeClose(ref bool Cancel)
		{
			Debug.WriteLine("Close");
		}

		private void App_WindowDeactivate(Excel.Workbook Wb, Excel.Window Wn)
		{
			Debug.WriteLine("Window Deactivate");
			Debug.WriteLine(Wb.Name);
		}

		public void Shutdown()
		{
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

			var hWndXLDesk = HandleUtility.GetXLDesk((IntPtr)workbook.Application.Hwnd);
			if (hWndXLDesk == IntPtr.Zero)
			{
				Debug.WriteLine("hWndXLDesk : null");
				return;
			}

			var hWndExcel7s = HandleUtility.GetExcel7(hWndXLDesk);
			if (hWndExcel7s != null)
			{
				//Debug.WriteLine("hWndExcel7s : " + hWndExcel7s.Count);

				foreach (var hWndExcel7 in hWndExcel7s)
				{
					this.CM.Excel7 = new Excel7Control(hWndExcel7);
					this.CM.Excel7.Changed += this.Excel7_Changed;
					break;
				}
			}
		}

		private void Unhook(Excel.Workbook workbook)
		{
			if (workbook != null)
			{
				workbook.SheetActivate -= this.Workbook_SheetActivate;
			}

			if (this.CM.Excel7 != null)
			{
				this.CM.Excel7.Changed -= this.Excel7_Changed;
				this.CM.Excel7.Dispose();
				this.CM.Excel7 = null;
			}

			this.Timer?.Dispose();
		}

		private void Workbook_SheetActivate(object sheet)
		{
			Debug.WriteLine("Activate");

			this.VMM.SheetListControl?.WorkbookInfo?.Select(sheet);
		}

		private void Excel7_Changed(object sender, EventArgs e)
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
