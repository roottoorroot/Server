using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Net;
using System.Net.Sockets;
using System.Threading;
using System.IO;
using System.Net.Mail;
using Excel = Microsoft.Office.Interop.Excel;
using System.Web;
using System.Net.Mime;

namespace Prototype_v1_Server
{
    public partial class Llog : Form
    {
        public string _patch; //Patch to logfile

        //Constructor for the show to new filelog from configFile;
        public Llog(string patch)
        {
            _patch = patch;
            InitializeComponent();
            InitText();
        }

        //Empty constructor
        public Llog()
        {
            InitializeComponent();
            InitText();
        }

        //Function forming mapping contents logfile in richTextbox on this form
        public void InitText()
        {
            try
            {
                StreamReader reader = File.OpenText(_patch);
                richTextBox1.Text = reader.ReadToEnd();
                reader.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Ошибка в InitText: \r" + e.ToString());
            }
        }
    }
}
