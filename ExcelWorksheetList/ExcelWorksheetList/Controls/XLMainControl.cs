using System;
using System.Windows.Forms;

namespace Toybox.ExcelWorksheetList.Controls
{
	using Utility;

	public class XLMainControl : NativeWindow, IDisposable
	{

		#region [IDisposable]

		public void Dispose()
		{
			this.ReleaseHandle();
		}

		#endregion [IDisposable]


		#region Static

		private readonly string MessageName = "DUserUnSubClassMessage";

		#endregion


		#region Constructor

		public XLMainControl()
		{
			this.MessageId = (int)User32Utility.RegisterWindowMessage(MessageName);
		}

		public XLMainControl(IntPtr hWnd)
			: this()
		{
			if (hWnd == IntPtr.Zero)
			{
				throw new ArgumentException("hWnd");
			}

			this.AssignHandle(hWnd);
		}

		~XLMainControl()
		{
			this.Dispose();
		}

		#endregion Constructor


		#region Event

		public event EventHandler Changed;

		#endregion Event


		#region Private Members

		private int MessageId { get; set; }

		#endregion Private Members


		#region Protected Methods

		protected override void WndProc(ref Message m)
		{
			base.WndProc(ref m);

			if (m.Msg == this.MessageId)
			{
				this.OnChanged();
			}
		}

		protected virtual void OnChanged()
		{
			this.Changed?.Invoke(this, EventArgs.Empty);
		}

		#endregion Protected Methods

	}
}
