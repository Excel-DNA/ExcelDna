using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

using ExcelDna.Integration;
using System.IO;

namespace ExcelDna.Logging
{

    internal partial class LogDisplayForm : Form
    {
        delegate void WriteLineDelegate(string messages);
        WriteLineDelegate writelineDelegate;

        IWin32Window _parentWindow;

        public LogDisplayForm(IWin32Window parentWindow)
        {
            _parentWindow = parentWindow;
            InitializeComponent();
            // Force handle creation so we can can BeginInvoke before ever having shown the window.
            CreateHandle();
            Text = DnaLibrary.CurrentLibraryName + " - Log Display";
            writelineDelegate = this.WriteLine;
        }


        void ShowLogDisplay()
        {
            Text = DnaLibrary.CurrentLibraryName + " - Log Display";
            if (!Visible)
            {
                CenterToParent();
                Show(_parentWindow);
            }
        }

        void HideLogDisplay()
        {
            // How can we prevent Excel from also being pushed to back?
            Hide();
        }

        private struct SyntaxHighlightingFormat
        {
            public SyntaxHighlightingFormat(string stringToMatch, Color color, FontStyle fontStyle)
            {
                StringToMatch = stringToMatch;
                Color = color;
                FontStyle = fontStyle;
            }
            readonly public string StringToMatch;
            readonly public Color Color;
            readonly public FontStyle FontStyle;
        }

        // When a format string is a substring of another, keep the latter first in the sequence!
        private static IEnumerable<SyntaxHighlightingFormat> syntxHighlights = new SyntaxHighlightingFormat[]
        {
            new SyntaxHighlightingFormat("INTERNAL ERROR", Color.Red, FontStyle.Bold),
            new SyntaxHighlightingFormat("ERROR", Color.Green, FontStyle.Bold)
        };

        static string lastMessage = null;

        private void WriteLine(string message)
        {
            // We don't want to emit the same error message twice. NB Remember all this happens single-threaded!
            if (string.IsNullOrEmpty(lastMessage) || lastMessage != message)
            {
                var oldLength = logMessages.Text.Length;
                // Set default formatting
                logMessages.SelectionStart = logMessages.Text.Length;
                logMessages.SelectionLength = 0;
                logMessages.SelectionColor = logMessages.ForeColor;
                logMessages.SelectionFont = logMessages.Font;
                logMessages.AppendText(message + "\r\n");
                lastMessage = message;
                foreach (var item in syntxHighlights)
                {
                    var stringToMatch = item.StringToMatch;
                    if (message.StartsWith(stringToMatch, true, CultureInfo.InvariantCulture))
                    {
                        // Highlight
                        logMessages.Focus();
                        logMessages.SelectionStart = oldLength;
                        logMessages.SelectionLength = stringToMatch.Length;
                        logMessages.SelectionColor = item.Color;
                        logMessages.SelectionFont = new Font(logMessages.Font, item.FontStyle);
                        break;
                    }
                }
                // Scroll to end and restore format defaults
                logMessages.SelectionStart = logMessages.Text.Length;
                logMessages.SelectionLength = 0;
                logMessages.ScrollToCaret();
            }
            ShowLogDisplay();
        }

        public void Clear()
        {
            if (InvokeRequired)
            {
                Invoke((MethodInvoker)(delegate() { logMessages.Clear(); }));
            }
            else
            {
                logMessages.Clear();
            }
        }

        public void WriteLine(string format, params object[] args)
        {
            try
            {
                var message = String.Format(format, args);
                if (InvokeRequired)
                {
                    Invoke(writelineDelegate, message);
                }
                else
                {
                    WriteLine(message);
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine("LogDisplayForm WriteLine failed: " + e);
            }
        }

        // When a form is closed, the handle gets disposed of. This means a successive
        // Show() call will recreate it in a potentially wrong thread. Esp. in 2007 this
        // can easily happen in one of the recalc threads. Those threads block and wait for
        // incoming calc requests, which means the form will block as well.
        // The solution is to only hide the form. 
        private void LogDisplayForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            HideLogDisplay();
            e.Cancel = true;
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
                File.WriteAllText(sfd.FileName, logMessages.Text);
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            logMessages.Clear();
        }

        private void LogDisplayForm_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Escape)
            {
                HideLogDisplay();
            }
        }

        private void logMessages_LinkClicked(object sender, LinkClickedEventArgs e)
        {
            System.Diagnostics.Process.Start(e.LinkText);
        }
    }

    public static class LogDisplay
    {
        static LogDisplayForm _form;

        internal static void CreateInstance()
        {
            if (_form == null)
            {
                _form = new LogDisplayForm(NativeWindow.FromHandle(ExcelDnaUtil.WindowHandle));
            }
        }

        [Obsolete("Rather use LogDisplay.Clear() and LogDisplay.WriteLine(...)")]
        public static void SetText(string text)
        {
            WriteLine(text);
        }

        public static void WriteLine(string format, params object[] args)
        {
            Debug.WriteLine("LogDisplay.WriteLine Start.");
            _form.WriteLine(format, args);
            Debug.WriteLine("LogDisplay.WriteLine Finished.");
        }

        public static void Clear()
        {
            _form.Clear();
        }
    }
}