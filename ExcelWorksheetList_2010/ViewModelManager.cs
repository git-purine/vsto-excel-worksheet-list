using Toybox.ExcelWorksheetList.ViewModels;

namespace ExcelWorksheetList_2010
{
	public class ViewModelManager
	{

		#region Constructor

		public ViewModelManager()
		{
			this.SheetListControl = new VM_SheetListControl();
		}

		#endregion Constructor


		#region Public Members

		public VM_SheetListControl SheetListControl { get; private set; }

		#endregion Public Members

	}
}
