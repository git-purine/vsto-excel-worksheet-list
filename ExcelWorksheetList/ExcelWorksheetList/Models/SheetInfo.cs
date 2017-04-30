using System.Windows;
using System.Windows.Media;
using System.Diagnostics;
using System.Windows.Media.Imaging;

namespace Toybox.ExcelWorksheetList.Models
{
	using Core.ComponentModel;
	using Extensions;

	public class SheetInfo : Model
	{

		#region Constructor

		public SheetInfo(object sheet)
		{
			this.Sheet = sheet;

			this.FontWeight = FontWeights.Normal;
			this.HeaderColor = Brushes.Transparent;
			this.ProtectionVisibility = Visibility.Collapsed;
			this.InvisibleVisibility = Visibility.Collapsed;

			this.Update();
		}

		#endregion Constructor


		#region Public Members

		public object Sheet
		{
			get; private set;
		}

		public string Name
		{
			get { return this.Sheet.InvokeMember(nameof(Name)) as string; }
		}

		public bool IsSelected
		{
			get { return this._isSelected; }
			set
			{
				if (this._isSelected == value) return;
				this._isSelected = value;

				this.FontWeight = this._isSelected ? FontWeights.Bold : FontWeights.Normal;

				if (this._isSelected)
				{
					this.Sheet.InvokeMethod("Activate");
				}

				this.OnPropertyChanged(nameof(IsSelected));
			}
		}
		private bool _isSelected;

		public FontWeight FontWeight
		{
			get { return this._fontWeight; }
			set
			{
				if (this._fontWeight == value) return;
				this._fontWeight = value;
				this.OnPropertyChanged(nameof(FontWeight));
			}
		}
		private FontWeight _fontWeight;

		public Brush HeaderColor
		{
			get { return this._headColor; }
			set
			{
				if (this._headColor == value) return;
				this._headColor = value;
				this.OnPropertyChanged(nameof(HeaderColor));
			}
		}
		private Brush _headColor;

		/// <summary>
		/// Protected
		/// </summary>
		public BitmapImage ProtectionIcon
		{
			get { return this._protectionIcon; }
			set
			{
				if (this._protectionIcon == value) return;
				this._protectionIcon = value;
				base.OnPropertyChanged(nameof(ProtectionIcon));
			}
		}
		private BitmapImage _protectionIcon;

		/// <summary>
		/// Protected
		/// </summary>
		public Visibility ProtectionVisibility
		{
			get { return this._protectionVisibility; }
			set
			{
				if (this._protectionVisibility == value) return;
				this._protectionVisibility = value;
				this.OnPropertyChanged(nameof(ProtectionVisibility));
			}
		}
		private Visibility _protectionVisibility;

		/// <summary>
		/// Invisible
		/// </summary>
		public Visibility InvisibleVisibility
		{
			get { return this._invisibleVisibility; }
			set
			{
				if (this._invisibleVisibility == value) return;
				this._invisibleVisibility = value;
				this.OnPropertyChanged(nameof(InvisibleVisibility));
			}
		}
		private Visibility _invisibleVisibility;

		//public Brush Background
		//{
		//	get { return this._background; }
		//	set
		//	{
		//		if (this._background == value) return;
		//		this._background = value;
		//		this.OnPropertyChanged(nameof(Background));
		//	}
		//}
		//private Brush _background;

		#endregion Public Members


		#region Public Methods

		public void Update()
		{
			// color
			var tab = this.Sheet.InvokeMember("Tab");
			var tabColor = tab.InvokeMember("Color");
			if (tabColor is bool && !(bool)tabColor)
			{
				this.HeaderColor = Brushes.Transparent;
			}
			else if (tabColor is int)
			{
				var value = ((int)tabColor);
				var r = value & 0xff;
				var g = (value >> 8) & 0xff;
				var b = (value >> 16) & 0xff;

				var color = Color.FromArgb(0xFF, (byte)r, (byte)g, (byte)b);

				this.HeaderColor = new SolidColorBrush(color);
			}

			// protection
			var protectionContents = this.Sheet.InvokeMember("ProtectContents");
			this.ProtectionVisibility = (protectionContents is bool && (bool)protectionContents) ? Visibility.Visible : Visibility.Collapsed;

			// invisible
			var visible = this.Sheet.InvokeMember("Visible");
			this.InvisibleVisibility = (visible is int && (int)visible == 0) ? Visibility.Visible : Visibility.Collapsed;

			this.OnPropertyChanged(nameof(Name));
		}

		#endregion Public Methods

	}
}
