using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;

namespace CallMeBack
{
    public partial class Form1 : Form
    {
        protected Microsoft.Office.Interop.Outlook.Application App;

        public Form1(Microsoft.Office.Interop.Outlook.Application _app)
        {
            App = _app;
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            
        }
		

    }
}
