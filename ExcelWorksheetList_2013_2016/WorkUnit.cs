using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Toybox.Core.ComponentModel;
using Toybox.ExcelWorksheetList.Controls;
using Toybox.ExcelWorksheetList.ViewModels;
using Toybox.Utility;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelWorksheetList_2013_2016
{
	public class WorkUnitCollection : List<WorkUnit>
	{

		public WorkUnit this[Excel.Workbook workbook]
		{
			get { return this.FirstOrDefault(unit => unit.Workbook == workbook); }
		}

	}
	
	public class WorkUnit : ViewModel, IDisposable
	{

		#region [IDisposable]

		public void Dispose()
		{
			this.Unhook();

			this.SheetListControl.Workbook = null;
			this.ContainerControl.Dispose();
		}

		#endregion [IDisposable]

		#region Constructor

		public WorkUnit(Excel.Workbook workbook)
		{
			this.Name = workbook.Name;
			this.Workbook = workbook;

			this.SheetListControl = new VM_SheetListControl();
			this.SheetListControl.Workbook = this.Workbook;

			this.ContainerControl = new WPFContainerControl();
			this.ContainerControl.DataContext = this.SheetListControl;

			this.Hook();
		}

		#endregion Constructor

		#region Public Members

		public string Name { get; protected set; }

		public Excel.Workbook Workbook { get; protected set; }

		public CustomTaskPane Pane { get; set; }

		public VM_SheetListControl SheetListControl { get; protected set; }

		public WPFContainerControl ContainerControl { get; protected set; }

		public Excel7Control Excel7Control { get; protected set; }

		#endregion Public Members

		#region Protected Members

		protected System.Threading.Timer Timer { get; set; }
		
		#endregion Protected Members


		#region Public Methods



		#endregion Public Methods

		#region Private Methods

		private void Hook()
		{
			if (this.Workbook == null)
			{
				return;
			}

			this.Workbook.SheetActivate += this.Workbook_SheetActivate;

			var hWndXLDesk = HandleUtility.GetXLDesk((IntPtr)this.Workbook.Application.Hwnd);
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
					this.Excel7Control = new Excel7Control(hWndExcel7);
					this.Excel7Control.Changed += this.Excel7_Changed;
					break;
				}
			}
		}

		private void Unhook()
		{
			if (this.Workbook != null)
			{
				this.Workbook.SheetActivate -= this.Workbook_SheetActivate;
			}

			if (this.Excel7Control != null)
			{
				this.Excel7Control.Changed -= this.Excel7_Changed;
				this.Excel7Control.Dispose();
				this.Excel7Control = null;
			}

			this.Timer?.Dispose();
		}

		private void Workbook_SheetActivate(object sheet)
		{
			Debug.WriteLine("Activate");

			this.SheetListControl?.WorkbookInfo?.Select(sheet);
		}

		private void Excel7_Changed(object sender, EventArgs e)
		{
			Debug.WriteLine("Changeing");

			this.Timer?.Dispose();
			this.Timer = new System.Threading.Timer((_) =>
			{
				Debug.WriteLine("Changed");

				this.ContainerControl?.Invoke((MethodInvoker)delegate ()
				{
					this.SheetListControl?.Update();
				});

			}, null, 50, System.Threading.Timeout.Infinite);
		}

		#endregion Private Methods

	}
}
