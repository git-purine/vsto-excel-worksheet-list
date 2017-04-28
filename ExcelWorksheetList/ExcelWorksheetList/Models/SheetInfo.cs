using System.Windows;
using System.Windows.Media;
using System.Diagnostics;

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
			this.VisibilityProtection = Visibility.Collapsed;
			this.VisibilityInvisible = Visibility.Collapsed;

			//this.Foreground = Brushes.Black;
			//this.Background = Brushes.Transparent;

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

		public Visibility VisibilityProtection
		{
			get { return this._visibilityProtection; }
			set
			{
				if (this._visibilityProtection == value) return;
				this._visibilityProtection = value;
				this.OnPropertyChanged(nameof(VisibilityProtection));
			}
		}
		private Visibility _visibilityProtection;

		public Visibility VisibilityInvisible
		{
			get { return this._visibilityInvisible; }
			set
			{
				if (this._visibilityInvisible == value) return;
				this._visibilityInvisible = value;
				this.OnPropertyChanged(nameof(VisibilityInvisible));
			}
		}
		private Visibility _visibilityInvisible;

		//public Brush Foreground
		//{
		//	get { return this._foreground; }
		//	set
		//	{
		//		if (this._foreground == value) return;
		//		this._foreground = value;
		//		this.OnPropertyChanged(nameof(Foreground));
		//	}
		//}
		//private Brush _foreground;

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
				//this.Background = Brushes.Transparent;
			}
			else if (tabColor is int)
			{
				var value = ((int)tabColor);
				var r = value & 0xff;
				var g = (value >> 8) & 0xff;
				var b = (value >> 16) & 0xff;

				var color = Color.FromArgb(0xFF, (byte)r, (byte)g, (byte)b);

				this.HeaderColor = new SolidColorBrush(color);
				//this.Foreground = color.GetBrightness() < 0.5 ? Brushes.White : Brushes.Black;
				//this.Background = new SolidColorBrush(color);
			}

			// protection
			var protectionContents = this.Sheet.InvokeMember("ProtectContents");
			this.VisibilityProtection = (protectionContents is bool && (bool)protectionContents) ? Visibility.Visible : Visibility.Collapsed;

			// invisible
			var visible = this.Sheet.InvokeMember("Visible");
			this.VisibilityInvisible = (visible is int && (int)visible == 0) ? Visibility.Visible : Visibility.Collapsed;

			this.OnPropertyChanged(nameof(Name));
		}

		#endregion Public Methods

	}
}
