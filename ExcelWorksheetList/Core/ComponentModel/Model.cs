using System;
using System.ComponentModel;

namespace Toybox.Core.ComponentModel
{
	/// <summary>
	/// Model
	/// </summary>
	[Serializable]
	public abstract class Model : IModel
	{

		#region [INotifyPropertyChanged]

		[field: NonSerialized]
		public event PropertyChangedEventHandler PropertyChanged;

		protected void OnPropertyChanged(string propertyName = null)
		{
			this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
		}

		#endregion [INotifyPropertyChanged]

	}
}