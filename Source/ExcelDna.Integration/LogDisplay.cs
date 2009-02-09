using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

using ExcelDna.Integration;

namespace ExcelDna.Logging
{

    public partial class LogDisplayForm : Form
    {

        public LogDisplayForm()
        {
            InitializeComponent();
            Text = DnaLibrary.CurrentLibraryName + " - Log Display";
        }

        public void SetText(string message)
        {
            listBoxErrors.Items.Clear();
            listBoxErrors.Items.Add(message);
        }

        public void AppendText(string message)
        {
            listBoxErrors.Items.Add(message);
            // Select last item ... and clear.
            listBoxErrors.SelectedItems.Clear();
            listBoxErrors.SelectedItem = listBoxErrors.Items[listBoxErrors.Items.Count - 1];
            listBoxErrors.SelectedItems.Clear();

        }
    }
    
    public class LogDisplay
    {
        static LogDisplayForm logDisplayForm;

        static public LogDisplayForm LogDisplayForm
        {
            get
            {
                if (logDisplayForm == null)
                {
                    logDisplayForm = new LogDisplayForm();
                    logDisplayForm.FormClosed += logDisplayForm_FormClosed;
                }
                return logDisplayForm;
            }
        }

        static void logDisplayForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            logDisplayForm = null;
        }

        public static void SetText(string message)
        {
            try
            {
                LogDisplayForm.SetText(message);
                if (!LogDisplayForm.Visible)
                    LogDisplayForm.Show( NativeWindow.FromHandle(ExcelDnaUtil.WindowHandle) );
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
        public static void WriteLine(string message)
        {
            try
            {

                LogDisplayForm.AppendText(message);
                if (!LogDisplayForm.Visible)
                    LogDisplayForm.Show(NativeWindow.FromHandle(ExcelDnaUtil.WindowHandle));
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
    }

}