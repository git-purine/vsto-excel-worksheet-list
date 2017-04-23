namespace Toybox.ExcelWorksheetList.Controls
{
	partial class WPFContainerControl
	{
		/// <summary> 
		/// 必要なデザイナー変数です。
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary> 
		/// 使用中のリソースをすべてクリーンアップします。
		/// </summary>
		/// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region コンポーネント デザイナーで生成されたコード

		/// <summary> 
		/// デザイナー サポートに必要なメソッドです。このメソッドの内容を 
		/// コード エディターで変更しないでください。
		/// </summary>
		private void InitializeComponent()
		{
			this.elementHost = new System.Windows.Forms.Integration.ElementHost();
			this.sheetListControl = new Toybox.ExcelWorksheetList.Views.SheetListControl();
			this.SuspendLayout();
			// 
			// elementHost
			// 
			this.elementHost.Dock = System.Windows.Forms.DockStyle.Fill;
			this.elementHost.Location = new System.Drawing.Point(0, 0);
			this.elementHost.Name = "elementHost";
			this.elementHost.Size = new System.Drawing.Size(150, 150);
			this.elementHost.TabIndex = 0;
			this.elementHost.Text = "elementHost1";
			this.elementHost.Child = this.sheetListControl;
			// 
			// WPFContainerControl
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.Controls.Add(this.elementHost);
			this.Name = "WPFContainerControl";
			this.ResumeLayout(false);

		}

		#endregion

		public System.Windows.Forms.Integration.ElementHost elementHost;
		public Views.SheetListControl sheetListControl;
	}
}
