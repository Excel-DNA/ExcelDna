using System;
using System.Threading;
using System.Threading.Tasks;
using ExcelDna.Integration;

namespace AsyncFunctions
{
    public class SyncContext
    {
        public static DateTime asyncTestSyncContext()
        {
            using (new ExcelSynchronizationContextInstaller())
            {
                Task.Factory.StartNew(() => Thread.Sleep(2000))
                  .ContinueWith(t =>
                      {
                          Console.Beep(); Console.Beep();
                          try
                          {
                              dynamic xlApp = ExcelDnaUtil.Application;
                              xlApp.Range["F1"].Value = "We have waited long enough.";
                          }
                          catch
                          {
                              Console.Beep(); Console.Beep(); Console.Beep();
                          }
                      }
                  ,TaskScheduler.FromCurrentSynchronizationContext() );
                return DateTime.Now;
            }
        }
    }


    public class ExcelSynchronizationContextInstaller : IDisposable
    {
        readonly SynchronizationContext _oldContext;

        public ExcelSynchronizationContextInstaller()
        {
            _oldContext = SynchronizationContext.Current;
            SynchronizationContext.SetSynchronizationContext(new ExcelSynchronizationContext());
        }

        public void Dispose()
        {
            SynchronizationContext.SetSynchronizationContext(_oldContext);
        }
    }
    
	
}
