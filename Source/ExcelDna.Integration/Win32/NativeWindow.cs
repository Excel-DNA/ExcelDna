#if !USE_WINDOWS_FORMS

namespace ExcelDna.Integration.Win32
{
    internal class NativeWindow
    {
        public System.IntPtr Handle { get; }

        public virtual void CreateHandle(CreateParams cp)
        {

        }

        public virtual void DestroyHandle()
        {

        }

        protected virtual void WndProc(ref Message m)
        {

        }
    }
}

#endif
