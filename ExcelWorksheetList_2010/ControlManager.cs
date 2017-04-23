using Toybox.ExcelWorksheetList.Controls;
using Toybox.ExcelWorksheetList.Ribbons;

namespace ExcelWorksheetList_2010
{
	public class ControlManager 
	{

		#region Constructor

		public ControlManager()
		{
			this.ContainerControl = new WPFContainerControl();
			this.Ribbon = new VisibilityRibbon();
		}

		#endregion Constructor


		#region Public Members

		public WPFContainerControl ContainerControl { get; private set; }
		public VisibilityRibbon Ribbon { get; private set; }

		//public XLMainControl XLMain { get; set; }

		#endregion Public Members

	}
}
