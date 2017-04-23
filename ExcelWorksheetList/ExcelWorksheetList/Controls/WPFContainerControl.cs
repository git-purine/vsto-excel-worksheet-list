using System.Windows.Forms;

namespace Toybox.ExcelWorksheetList.Controls
{
	public partial class WPFContainerControl : UserControl
	{

		#region Constructor

		public WPFContainerControl()
		{
			InitializeComponent();
		}

		#endregion Constructor


		#region Public Members

		public object DataContext
		{
			get { return this.sheetListControl.DataContext; }
			set { this.sheetListControl.DataContext = value; }
		}

		#endregion Public Members

	}
}
