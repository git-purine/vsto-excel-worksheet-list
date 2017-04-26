using System.Diagnostics;
using Microsoft.Office.Tools;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace ExcelWorksheetList_2013_2016
{
	using Ribbons;

	public class AppManager
	{

		#region [Static]

		private static readonly string TITLE = "Sheet List";

		public static AppManager Instance { get; } = new AppManager();

		#endregion [Static]


		#region Constructor

		protected AppManager()
		{
			this.Ribbon = new WorksheetListRibbon();
			this.WorkUnits = new WorkUnitCollection();
		}

		#endregion Constructor


		#region Public Members

		public Office.IRibbonExtensibility Ribbon { get; protected set; }

		#endregion Public Members

		#region Protected Members

		protected Excel.Application App { get; set; }
		protected CustomTaskPaneCollection Panes { get; set; }
		protected WorkUnitCollection WorkUnits { get; set; }

		#endregion Protected Members


		#region Public Methods

		public void Startup(Excel.Application app, CustomTaskPaneCollection panes)
		{
			this.App = app;
			this.Panes = panes;

			// Active Hook
			this.App.WindowActivate += App_WindowActivate;
		}

		public void Shutdown()
		{
			// Dispose
			this.WorkUnits.ForEach(workUnit => workUnit.Dispose());
			WorkUnits.Clear();

			// Active Hook
			this.App.WindowActivate -= App_WindowActivate;
		}

		public void ShowActiveWorksheetPane()
		{
			var workUnit = this.WorkUnits[this.App.ActiveWorkbook];
			if(workUnit != null)
			{
				workUnit.Pane.Visible = true;
			}
		}

		#endregion Public Methods

		#region Private Methods

		private void App_WindowActivate(Excel.Workbook workbook, Excel.Window window)
		{
			Debug.WriteLine("App_WindowActivate");

			this.WorkUnits.ForEach((workUnit) =>
			{
				foreach (Excel.Workbook appWorkbook in this.App.Workbooks)
				{
					if (appWorkbook == workUnit.Workbook)
					{
						return;
					}
				}

				this.WorkUnits.Remove(workUnit);
				workUnit.Dispose();
			});

			if (this.WorkUnits[workbook] == null)
			{
				var workUnit = new WorkUnit(workbook);
				workUnit.Pane = this.Panes.Add(workUnit.ContainerControl, TITLE, window);
				workUnit.Pane.DockPosition = Office.MsoCTPDockPosition.msoCTPDockPositionLeft;
				workUnit.Pane.Visible = true;

				this.WorkUnits.Add(workUnit);
			}
		}

		#endregion Private Methods

	}
}
