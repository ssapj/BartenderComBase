using BarTender;
using System;
using System.Diagnostics;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace ssapj.BartenderComBase
{
	public class BartenderComBase : IDisposable
	{
		private Application _bartenderApplication;
		private readonly Task _initializationTask;
		private readonly bool _runAsync;
		private readonly CancellationToken _token;
		private bool _isInitializing;

		protected BartenderComBase(bool runAsync = false, CancellationToken token = default)
		{
			this._runAsync = runAsync;

			if (runAsync)
			{
				this._token = token;
				this._initializationTask = this.StartBartenderAsync(this._token);
			}
			else
			{
				this.StartBartender();
			}
		}

		private void StartBartender()
		{
			this._isInitializing = true;
			this._bartenderApplication = new Application();
			this._isInitializing = false;
		}

		private async Task StartBartenderAsync(CancellationToken token = default)
		{
			try
			{
				await Task.Run(() =>
				{
					try
					{
						token.ThrowIfCancellationRequested();
						this.StartBartender();
					}
					catch
					{
						throw;
					}
				}, token).ConfigureAwait(false);
			}
			catch
			{
				throw;
			}
		}

		protected async Task<Application> GetBartenderApplicationAsync()
		{
			this._token.ThrowIfCancellationRequested();

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
				if (this._runAsync && !this._token.IsCancellationRequested)
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
							break;
						default:
							break;
					}

					this._initializationTask.Dispose();
				}

			}

			while (this._isInitializing)
			{
				Task.Delay(10);
			}

			if (this._bartenderApplication != null)
			{
				var processes = Process.GetProcesses();

				if (processes.Any(x => x.Id == this._bartenderApplication.ProcessId))
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