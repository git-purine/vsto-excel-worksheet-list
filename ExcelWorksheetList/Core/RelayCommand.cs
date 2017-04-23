using System;
using System.Diagnostics;
using System.Windows.Input;

namespace Toybox.Core.Windows.Input
{
	/// <summary>
	/// RelayCommand
	/// </summary>
	public class RelayCommand : RelayCommand<object>
	{

		#region Constructor

		/// <summary>
		/// RelayCommand
		/// </summary>
		/// <param name="execute"></param>
		/// <param name="canExecute"></param>
		public RelayCommand(Action<object> execute, Predicate<object> canExecute = null)
			: base(execute, canExecute)
		{
		}

		#endregion Constructor

	}

	/// <summary>
	/// RelayCommand
	/// </summary>
	/// <typeparam name="T"></typeparam>
	public class RelayCommand<T> : ICommand
	{

		#region [ICommand]

		/// <summary>
		/// 実行検証メソッド
		/// </summary>
		protected readonly Predicate<T> _canExecute;

		/// <summary>
		/// 実行メソッド
		/// </summary>
		protected readonly Action<T> _execute;

		/// <summary>
		/// 実行検証メソッド変更
		/// </summary>
		public event EventHandler CanExecuteChanged
		{
			add { CommandManager.RequerySuggested += value; }
			remove { CommandManager.RequerySuggested -= value; }
		}

		/// <summary>
		/// 実行検証
		/// </summary>
		/// <param name="param"></param>
		/// <returns></returns>
		[DebuggerStepThrough]
		public bool CanExecute(object param)
		{
			return this._canExecute == null ? true : this._canExecute((T)param);
		}

		/// <summary>
		/// 実行
		/// </summary>
		/// <param name="param"></param>
		public void Execute(object param)
		{
			this._execute((T)param);
		}

		#endregion [ICommand]


		#region Constructor

		/// <summary>
		/// RelayCommand
		/// </summary>
		/// <param name="execute"></param>
		/// <param name="canExecute"></param>
		public RelayCommand(Action<T> execute, Predicate<T> canExecute = null)
		{
			if (execute == null)
				throw new ArgumentNullException(nameof(execute));

			this._execute = execute;
			this._canExecute = canExecute;
		}

		#endregion Constructor

	}
}
