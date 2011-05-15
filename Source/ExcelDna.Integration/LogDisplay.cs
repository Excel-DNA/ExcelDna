using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Windows.Forms;

using ExcelDna.Integration;

namespace ExcelDna.Logging
{

    internal partial class LogDisplayForm : Form
    {
        [DllImport("user32.dll")]
        static extern IntPtr SetFocus(IntPtr hWnd);
        internal static LogDisplayForm _form;

        internal static void ShowForm()
        {
            if (_form == null)
            {
                _form = new LogDisplayForm();
            }
            
            if (_form.Visible == false)
            {
                _form.Show(null);
                // SetFocus(ExcelDnaUtil.WindowHandle);
            }
        }

        internal static void HideForm()
        {
            if (_form != null)
            {
                _form.updateTimer.Enabled = false;
                _form.Close();
            }
        }

        System.Windows.Forms.Timer updateTimer;

        internal LogDisplayForm()
        {
            InitializeComponent();
            Text = DnaLibrary.CurrentLibraryName + " - Log Display";
            CenterToParent();
            logMessages.VirtualListSize = LogDisplay.LogStrings.Count;
            updateTimer = new System.Windows.Forms.Timer();
            updateTimer.Interval = 250;
            updateTimer.Tick += updateTimer_Tick;
            updateTimer.Enabled = true;
        }

        void updateTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                if (!IsDisposed && updateTimer.Enabled && LogDisplay.LogStringsUpdated)
                {
                    // CONSIDER: There are some race conditions here 
                    // - but I'd rather have some log mis-painting than deadlock between the UI thread and a calculation thread.
                    logMessages.VirtualListSize = LogDisplay.LogStrings.Count;
                    LogDisplay.LogStringsUpdated = false;
                    //ClearCache();
                    logMessages.Invalidate();
                    logMessages.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent);
                    // Debug.Print("LogDisplayForm.updateTimer_Tick - Updated to " + logMessages.VirtualListSize + " messages");
                }
            }
            catch (Exception ex)
            {
                Debug.Print("Exception in updateTime_Tick: " + ex);
            }
        }

        private ListViewItem MakeItem(int messageIndex)
        {
            string message;
            try
            {
                message = LogDisplay.LogStrings[messageIndex];
                if (message.Length > 259)
                {
                    message = message.Substring(0, 253) + " [...]";
                }
            }
            catch
            {
                message = " ";
            }

            return new ListViewItem(message);
        }

        public void Clear()
        {
            logMessages.VirtualListSize = 0;
            logMessages.AutoResizeColumn(0, ColumnHeaderAutoResizeStyle.ColumnContent);
        }

        private void LogDisplayForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            _form.updateTimer.Enabled = false;
            _form = null;
            LogDisplay.IsFormVisible = false;
            try
            {
                SetFocus(ExcelDnaUtil.WindowHandle);
            }
            catch { }   // Probably not in Excel !?
        }

        private void btnSaveErrors_Click(object sender, EventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.DefaultExt = "txt";
            sfd.Filter = "Text File (*.txt)|*.txt|All Files (*.*)|*.*";
            sfd.Title = "Save Error List As";
            DialogResult result = sfd.ShowDialog();
            if (result == DialogResult.OK)
            {
                File.WriteAllText(sfd.FileName, LogDisplay.GetAllText());
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            LogDisplay.Clear();
            Clear();
        }

        private void LogDisplayForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                Close();
            }
        }

        private void btnCopy_Click(object sender, EventArgs e)
        {
            string allText = LogDisplay.GetAllText();
            if (allText != null)
            {
                Clipboard.SetText(allText);
            }
        }

        private void logMessages_RetrieveVirtualItem(object sender, RetrieveVirtualItemEventArgs e)
        {
            e.Item = MakeItem(e.ItemIndex);
        }
    }

    public class MessageBufferChangedEventArgs : EventArgs
    {
        public enum MessageBufferChangeType
        {
            Clear,
            RemoveFirst,
            AddLast
        }

        public string AddedMessage { get; set; }

        public MessageBufferChangeType ChangeType { get; set; }

        public MessageBufferChangedEventArgs(MessageBufferChangeType changeType)
        {
            ChangeType = changeType;
        }

        public MessageBufferChangedEventArgs(string addedMessage)
        {
            ChangeType = MessageBufferChangeType.AddLast;
            AddedMessage = addedMessage;
        }
    }

    public static class LogDisplay
    {
        internal static List<string> LogStrings;
        internal static bool LogStringsUpdated;
        const int maxLogSize = 10000;   // Max number of strings we'll allow in the buffer. Individual strings are unbounded.
        internal static object SyncRoot = new object();
        internal static bool IsFormVisible;

        static SynchronizationContext _syncContext;

        // This must be called on the main Excel thread.
        internal static void CreateInstance()
        {
            LogStrings = new List<string>();
            LogStringsUpdated = true;
            _syncContext = SynchronizationContext.Current;
            IsFormVisible = false;
            if (_syncContext == null)
            {
                _syncContext = new WindowsFormsSynchronizationContext();
                //Debug.Print("LogDisplay.CreateInstance - Creating SyncContext on thread: " + Thread.CurrentThread.ManagedThreadId);
            }
        }

        public static void Show()
        {
           _syncContext.Post(delegate(object state)
                {
                    LogDisplayForm.ShowForm();
                }, null);
        }

        public static void Hide()
        {
            _syncContext.Post(delegate(object state)
            {
                LogDisplayForm.HideForm();
            }, null);
        }

        [Obsolete("Rather use LogDisplay.Clear() and LogDisplay.WriteLine(...)")]
        public static void SetText(string text)
        {
            WriteLine(text);
        }

        // This might be called from any calculation thread - also displays the form.
        public static void WriteLine(string format, params object[] args)
        {
            //Debug.WriteLine("LogDisplay.WriteLine start on thread " + System.Threading.Thread.CurrentThread.ManagedThreadId);
            lock (SyncRoot)
            {
                if (!IsFormVisible)
                {
                    Show();
                    IsFormVisible = true;
                }
                RecordLine(format, args);
            }
            //Debug.WriteLine("LogDisplay.WriteLine completed in thread " + System.Threading.Thread.CurrentThread.ManagedThreadId);
        }

        // This might be called from any calculation thread
        // Does not force a Show
        public static void RecordLine(string format, params object[] args)
        {
            lock (SyncRoot)
            {
                if (LogStrings.Count > maxLogSize)
                {
                    LogStrings.RemoveAt(0);
                }
                string message = string.Format(format, args);
                string[] messageLines = message.Split(new char[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
                LogStrings.AddRange(messageLines);
                LogStringsUpdated = true;
            }
        }

        public static void Clear()
        {
            lock (SyncRoot)
            {
                LogStrings.Clear();
                LogStringsUpdated = true;
            }
        }

        internal static string GetAllText()
        {
            StringBuilder sb = new StringBuilder();
            foreach (string msg in LogStrings)
            {
                sb.AppendLine(msg);
            }
            return sb.ToString();
        }

    }
}