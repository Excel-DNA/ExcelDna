//  Copyright (c) Govert van Drimmelen. All rights reserved.
//  Excel-DNA is licensed under the zlib license. See LICENSE.txt for details.

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using ExcelDna.Integration;

namespace ExcelDna.Logging
{
    internal partial class LogDisplayForm
    {
        [DllImport("user32.dll")]
        static extern IntPtr SetFocus(IntPtr hWnd);
        internal static LogDisplayForm _form;

        internal static void ShowForm()
        {
        }

        internal static void HideForm()
        {
        }

        internal LogDisplayForm()
        {
        }

        public void Clear()
        {
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

    public enum DisplayOrder
    {
        NewestLast,
        NewestFirst
    }

    public static class LogDisplay
    {
        internal static List<string> LogStrings = new List<string>();
        internal static bool LogStringsUpdated;
        const int maxLogSize = 10000;   // Max number of strings we'll allow in the buffer. Individual strings are unbounded.
        internal static object SyncRoot = new object();
        internal static bool IsFormVisible;

        // _syncContext is null until we call CreateInstance, which 
        static SynchronizationContext _syncContext;

        // This must be called on the main Excel thread.
        internal static void CreateInstance()
        {
            LogStringsUpdated = true;
            _syncContext = SynchronizationContext.Current;
            IsFormVisible = false;
            if (_syncContext == null)
            {
                //_syncContext = new WindowsFormsSynchronizationContext();
                //Debug.Print("LogDisplay.CreateInstance - Creating SyncContext on thread: " + Thread.CurrentThread.ManagedThreadId);
            }
        }

        public static void Show()
        {
        }

        public static void Hide()
        {
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
                Show();
                RecordLine(format, args);
            }
            //Debug.WriteLine("LogDisplay.WriteLine completed in thread " + System.Threading.Thread.CurrentThread.ManagedThreadId);
        }

        static readonly char[] LineEndChars = new char[] { '\r', '\n' };
        // This might be called from any calculation thread
        // Does not force a Show
        public static void RecordLine(string format, params object[] args)
        {
            lock (SyncRoot)
            {
                string message = args.Length > 0 ? string.Format(format, args) : format;
                string[] messageLines = message.Split(LineEndChars, StringSplitOptions.RemoveEmptyEntries);
                if (DisplayOrder == DisplayOrder.NewestLast)
                {
                    // Insert at the end
                    LogStrings.AddRange(messageLines);
                }
                else
                {
                    // Insert at the beginning.
                    LogStrings.InsertRange(0, messageLines);
                }
                TruncateLog();
                LogStringsUpdated = true;
            }
        }

        static void TruncateLog()
        {
            while (LogStrings.Count > maxLogSize)
            {
                if (DisplayOrder == DisplayOrder.NewestLast)
                {
                    // Remove from the front
                    LogStrings.RemoveAt(0);
                }
                else
                {
                    // Remove from the back
                    LogStrings.RemoveAt(LogStrings.Count - 1);
                }
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

        static DisplayOrder _displayOrder;
        public static DisplayOrder DisplayOrder
        {
            get { return _displayOrder; }
            set
            {
                if (_displayOrder != value)
                {
                    _displayOrder = value;
                    LogStrings.Reverse();
                    LogStringsUpdated = true;
                }
            }
        }
    }
}
