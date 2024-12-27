using BarTender;
using System;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;

namespace ssapj.BartenderComBase
{
	public partial class BartenderComBase : IDisposable
	{
		private Application _bartenderApplication;
		private int _bartenderProcessId;

		protected Application BartenderInstance
		{
			get
			{
				if (this._bartenderApplication is null)
				{
					StartBartender();
				}

				return this._bartenderApplication;
			}
		}

		private void StartBartender()
		{
			this._bartenderApplication ??= new Application();
			this._bartenderProcessId = this._bartenderApplication.ProcessId;
		}

		#region IDisposable Support
		private bool _disposedValue;

		protected virtual void Dispose(bool disposing)
		{
			if (this._disposedValue)
			{
				return;
			}

			if (disposing)
			{
			}

			if (this._bartenderApplication != null)
			{
				var processes = Process.GetProcesses();

				if (processes.Any(x => x.Id == this._bartenderProcessId))
				{
					this._bartenderApplication.Quit(BtSaveOptions.btDoNotSaveChanges);
				}

				_ = Marshal.FinalReleaseComObject(this._bartenderApplication);
				this._bartenderApplication = null;
			}

			this._disposedValue = true;
		}

		~BartenderComBase()
		{
			this.Dispose(false);
		}

		public void Dispose()
		{
			this.Dispose(true);
			GC.SuppressFinalize(this);
		}
		#endregion
	}

}
