using BarTender;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace ssapj.BartenderComBase
{
	public class BartenderComBase : IDisposable
	{
		protected Application BartenderApplication;
		protected int ProcessIdOfBartenderApplication;
		protected Task InitializationTask;

		protected BartenderComBase(bool runAsync = false)
		{
			if (runAsync)
			{
				this.InitializationTask = this.StartBartenderAsync();
			}
			else
			{
				this.StartBartender();
			}
		}

		private void StartBartender()
		{
			this.BartenderApplication = new Application();
			this.ProcessIdOfBartenderApplication = this.BartenderApplication.ProcessId;
		}

		private async Task StartBartenderAsync()
		{
			await Task.Run(this.StartBartender).ConfigureAwait(false);
		}

		#region IDisposable Support
		private bool _disposedValue;

		protected virtual void Dispose(bool disposing)
		{
			if (this._disposedValue)
			{
				return;
			}

			//wait till BarTender wake up.
			this.InitializationTask.Wait();

			if (this.BartenderApplication != null)
			{
				try
				{
					using (Process.GetProcessById(this.ProcessIdOfBartenderApplication))
					{
						this.BartenderApplication.Quit(BtSaveOptions.btDoNotSaveChanges);
					}
				}
				finally
				{
					System.Runtime.InteropServices.Marshal.FinalReleaseComObject(this.BartenderApplication);
					this.BartenderApplication = null;
				}
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
