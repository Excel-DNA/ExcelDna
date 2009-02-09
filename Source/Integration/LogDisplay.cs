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
            Text = DnaLibrary.CurrentLibrary.Name + " - Log Display";
        }

        public void SetText(string message)
        {
            textBoxMessage.Text = message;
            textBoxMessage.Select(0, 0);
        }

        public void AppendText(string message)
        {
            textBoxMessage.Text += message;
            textBoxMessage.Select(textBoxMessage.Text.Length, 0);
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
                LogDisplayForm.Show();
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
    }

        public static void Write(string message)
        {
            try
            {
                LogDisplayForm.AppendText(message);
                LogDisplayForm.Show();
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

                LogDisplayForm.AppendText(message + "\r\n");
                LogDisplayForm.Show();
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }
    }

}