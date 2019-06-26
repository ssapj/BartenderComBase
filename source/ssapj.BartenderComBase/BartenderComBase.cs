using BarTender;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace ssapj.BartenderComBase
{
    public class BarTenderComBase : IDisposable
    {
        protected Application _bartenderApplication;
        protected int _processIdOfBartenderApplication;
        protected ValueTask _initializationValueTask;

        protected BarTenderComBase()
        {
            this.StartBartender();
        }

        protected BarTenderComBase(bool runAsync)
        {
            if (runAsync)
            {
                this._initializationValueTask = this.StartBartenderAsync();
            }
            else
            {
                this.StartBartender();
            }
        }

        private void StartBartender()
        {
            this._bartenderApplication = new Application();
            this._processIdOfBartenderApplication = this._bartenderApplication.ProcessId;
        }

        private async ValueTask StartBartenderAsync()
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

            if (this._bartenderApplication != null)
            {
                try
                {
                    using (Process.GetProcessById(this._processIdOfBartenderApplication))
                    {
                        this._bartenderApplication.Quit(BtSaveOptions.btDoNotSaveChanges);
                    }
                }
                finally
                {

                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(this._bartenderApplication);
                    this._bartenderApplication = null;

                }
            }

            this._disposedValue = true;
        }

        ~BarTenderComBase()
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
