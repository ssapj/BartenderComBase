using BarTender;
using System;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;

namespace ssapj.BartenderComBase
{
    public class BartenderComBase : IDisposable
    {
		private Application _bartenderApplication;
		private readonly Task _initializationTask;
		private readonly bool _runAsync;

		protected BartenderComBase(bool runAsync = false)
		{
			this._runAsync = runAsync;

			if (runAsync)
			{
				this._initializationTask = this.StartBartenderAsync();
			}
			else
			{
				this.StartBartender();
			}
		}

		private void StartBartender()
		{
			this._bartenderApplication = new Application();
		}

		private async Task StartBartenderAsync()
		{
			await Task.Run(this.StartBartender).ConfigureAwait(false);
		}

		protected async ValueTask<Application> GetBartenderApplicationAsync()
		{
			if (!this._runAsync || this._initializationTask.IsCompleted)
			{
				return this._bartenderApplication;
			}

			switch (this._initializationTask.Status)
			{
				case TaskStatus.Created:
				case TaskStatus.WaitingForActivation:
				case TaskStatus.WaitingToRun:
				case TaskStatus.Running:
				case TaskStatus.WaitingForChildrenToComplete:
					await this._initializationTask.ConfigureAwait(false);
					break;
				case TaskStatus.RanToCompletion:
					break;
				case TaskStatus.Canceled:
					throw new TaskCanceledException();
				case TaskStatus.Faulted:
					throw new Exception();
				default:
					break;
			}

			return this._bartenderApplication;
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
				if (this._runAsync)
				{
					//wait till BarTender wake up.
					switch (this._initializationTask.Status)
					{
						case TaskStatus.Created:
						case TaskStatus.WaitingForActivation:
						case TaskStatus.WaitingToRun:
						case TaskStatus.Running:
						case TaskStatus.WaitingForChildrenToComplete:
							this._initializationTask.Wait();
							break;
						case TaskStatus.RanToCompletion:
						case TaskStatus.Canceled:
						case TaskStatus.Faulted:
						default:
							break;
					}

					this._initializationTask.Dispose();
				}

			}

			if (this._bartenderApplication != null)
			{
				var processes = Process.GetProcesses();

				if (processes.Any(x => x.Id == this._bartenderApplication.ProcessId))
				{
					this._bartenderApplication.Quit(BtSaveOptions.btDoNotSaveChanges);
				}

				System.Runtime.InteropServices.Marshal.FinalReleaseComObject(this._bartenderApplication);
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
