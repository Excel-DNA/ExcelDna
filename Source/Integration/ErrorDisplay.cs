using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace ExcelDna.Integration
{
    public partial class ErrorDisplay : Form
    {
        public ErrorDisplay()
        {
            InitializeComponent();
        }

        public static void DisplayErrorMessage(string message)
        {
            ErrorDisplay form = new ErrorDisplay();
            form.textBoxMessage.Text = message;
            form.textBoxMessage.Select(0,0);
            form.Show();
        }
    }
}