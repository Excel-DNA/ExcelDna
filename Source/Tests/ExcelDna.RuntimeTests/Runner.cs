using System.Runtime.InteropServices;

namespace ExcelDna.RuntimeTests
{
    internal class Runner
    {
        public static void ExecuteWithRetryWhenExcelBusy(Action action, Action? retryAction = null)
        {
            for (int i = 0; i < 3; ++i)
            {
                try
                {
                    action();
                    return;
                }
                catch (COMException e) when (e.ErrorCode == -2147417846) // 0x8001010A (RPC_E_SERVERCALL_RETRYLATER)
                {
                    retryAction?.Invoke();
                }
            }
        }
    }
}
