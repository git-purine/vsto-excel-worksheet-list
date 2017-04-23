using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using Excel = Microsoft.Office.Interop.Excel;

namespace Toybox.ExcelWorksheetList.Models
{
	using Extensions;

	public class WorkbookInfo : ObservableCollection<SheetInfo>
	{

		#region Constructor

		public WorkbookInfo(Excel.Workbook workbook)
		{
			this.Workbook = workbook;
			this.Initialize();
		}

		#endregion Constructor


		#region Public Members

		public Excel.Workbook Workbook { get; private set; }

		#endregion Public Members


		#region Public Methods

		public void Update()
		{
			if (this.Workbook == null)
			{
				return;
			}

			Debug.WriteLine("Update Start");

			this.UpdateRemovedItems();

			this.UpdateAddedItems();

			this.UpdateSortedItems();

			this.UpdateAllItems();

			Debug.WriteLine("Update End");
		}

		public void Select(object sheet)
		{
			foreach (var wi in this)
			{
				wi.IsSelected = (wi.Sheet == sheet);
			}
		}

		#endregion Public Methods

		#region Private Methods

		private void Initialize()
		{
			if (this.Workbook?.Sheets == null)
			{
				return;
			}

			foreach (var sheet in this.Workbook.Sheets)
			{
				this.Add(new SheetInfo(sheet));
			}
		}

		private void UpdateAddedItems()
		{
			foreach (var sheet in this.Workbook.Sheets)
			{
				var isFind = false;

				foreach (var sheetInfo in this)
				{
					if (sheetInfo.Sheet == sheet)
					{
						isFind = true;
						break;
					}
				}

				if (isFind)
				{
					continue;
				}

				// 追加
				this.Add(new SheetInfo(sheet));
			}
		}

		private void UpdateRemovedItems()
		{
			for (var n = 0; n < this.Count;)
			{
				var sheetInfo = this[n];

				var isFind = false;

				foreach (var sheet in this.Workbook.Sheets)
				{
					if (sheetInfo.Sheet == sheet)
					{
						isFind = true;
						break;
					}
				}

				if (isFind)
				{
					n++;
					continue;
				}

				// 削除
				this.Remove(sheetInfo);
			}
		}

		private void UpdateSortedItems()
		{
			Func<object, Tuple<int, SheetInfo>> find = (_) =>
			{
				for (var n = 0; n < this.Count; n++)
				{
					var si = this[n];

					//if (_.Equals(si.Sheet))
					if (si.Name == _.InvokeMember("Name") as string)
					{
						return new Tuple<int, SheetInfo>(n, si);
					}
				}

				return null;
			};

			var index = 0;
			foreach (var sheet in this.Workbook.Sheets)
			{
				var tuple = find(sheet);
				if (tuple != null && tuple.Item1 != index)
				{
					this.Remove(tuple.Item2);
					this.Insert(index, tuple.Item2);
				}

				index++;
			}
		}

		private void UpdateAllItems()
		{
			foreach (var sheetInfo in this)
			{
				sheetInfo.IsSelected = sheetInfo.Sheet == this.Workbook.ActiveSheet;
				sheetInfo.Update();
			}
		}

		#endregion Private Methods

	}
}
