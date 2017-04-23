using System;

namespace Toybox.Core
{
	public class EventArgs<T> : EventArgs
	{

		#region Constructor

		public EventArgs(T obj)
		{
			this.Item = obj;
		}

		#endregion Constructor


		#region Public Members

		public T Item { get; protected set; }

		#endregion Public Members

	}
}
