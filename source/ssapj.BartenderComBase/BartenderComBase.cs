using BarTender;
using System;
using System.Diagnostics;
using System.Threading.Tasks;

namespace ssapj.BartenderComBase
{
    public class BarTenderComBase : IDisposable
    {
        protected Application BartenderApplication;
        protected int ProcessIdOfBartenderApplication;
        protected ValueTask InitializationValueTask;

        protected BarTenderComBase(bool runAsync = false)
        {
            if (runAsync)
            {
                this.InitializationValueTask = this.StartBartenderAsync();
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
