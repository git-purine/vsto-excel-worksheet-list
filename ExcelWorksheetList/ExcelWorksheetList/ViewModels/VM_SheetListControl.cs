using System.Collections.ObjectModel;
using System.Linq;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Input;

namespace Toybox.ExcelWorksheetList.ViewModels
{
	using Core.ComponentModel;
	using Core.Windows.Input;
	using Models;

	public class VM_SheetListControl : ViewModel
	{

		#region Public Members

		public Excel.Workbook Workbook
		{
			get { return this._workbook; }
			set
			{
				if (this._workbook == value) return;
				this._workbook = value;
				this.WorkbookInfo = new WorkbookInfo(value);
			}
		}
		private Excel.Workbook _workbook;

		public WorkbookInfo WorkbookInfo
		{
			get { return this._workbookInfo; }
			protected set
			{
				if (this._workbookInfo == value) return;
				this._workbookInfo = value;
				this.OnPropertyChanged(nameof(WorkbookInfo));
				this.OnPropertyChanged(nameof(FilteredWorkbookInfo));
			}
		}
		private WorkbookInfo _workbookInfo;

		public SheetInfo SelectedSheet
		{
			get { return this._selectedSheet; }
			set
			{
				if (this._selectedSheet == value) return;
				this._selectedSheet = value;
				if(this._selectedSheet != null)
				{
					this._selectedSheet.IsSelected = true;
				}
				this.OnPropertyChanged(nameof(SelectedSheet));
			}
		}
		private SheetInfo _selectedSheet;

		public string FilterText
		{
			get { return this._filterText; }
			set
			{
				if (this._filterText == value) return;
				this._filterText = value;
				this.OnPropertyChanged(nameof(FilterText));
				this.OnPropertyChanged(nameof(FilteredWorkbookInfo));
			}
		}
		private string _filterText;

		public ObservableCollection<SheetInfo> FilteredWorkbookInfo
		{
			get
			{
				var filtered = this.WorkbookInfo?.Where(si =>
				{
					return string.IsNullOrEmpty(this.FilterText)
							|| si.Name.ToLower().Contains(this.FilterText.Trim().ToLower());
				});

				return filtered == null ? new ObservableCollection<SheetInfo>() 
					                      : new ObservableCollection<SheetInfo>(filtered);
			}
		}

		#endregion Public Members


		#region Public Methods

		public void Update()
		{
			this.WorkbookInfo.Update();
			this.OnPropertyChanged(nameof(FilteredWorkbookInfo));
		}

		#endregion Public Methods


		#region Commands

		public ICommand Reload
		{
			get
			{
				return this._reload ?? (this._reload = new RelayCommand<object>((_) =>
				{
					this.Update();
				}));
			}
		}
		private ICommand _reload;

		#endregion Commands

	}
}
