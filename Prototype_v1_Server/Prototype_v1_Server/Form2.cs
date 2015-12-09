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
using System.Web;
using Microsoft.Office.Interop.Excel;
using System.Net.Mime;

namespace Prototype_v1_Server
{
    public partial class Form2 : Form
    {
        public string[] ttemp;

        public Form2()
        {
            InitializeComponent();
        }

        

        private void button1_Click(object sender, EventArgs e)
        {

            Form1 f1 = new Form1();
            
            string _patch = @"C:\log\Closelog.txt";
            DialogResult dl = MessageBox.Show("Закрыть отчётный периуд с " + dateTimePicker1.Text + " по: " + datePicker2.Text, "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dl == DialogResult.Yes)
            {
                string fileName = String.Empty;

                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                saveFileDialog1.Filter = "xls files (*.xls)|*.xls|All files (*.*)|*.*";
                saveFileDialog1.FilterIndex = 1;
                saveFileDialog1.RestoreDirectory = true;
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    fileName = saveFileDialog1.FileName;
                    //NEW
                     f1.toExcellFile(2, _patch, fileName);
                    //saveDataInExcel(dataGridView1,fileName);
                    MessageBox.Show("Отчётный период закрыт", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    FileInfo fi = new FileInfo(_patch);
                    FileStream flstr = fi.Create();
                    flstr.Close();
                   
                }
                else
                {
                    MessageBox.Show("__________");
                    //return;
                    ////сохраняем Workbook

                    //wb.SaveAs(fileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    //saveFileDialog1.Dispose();
                }
            }
        }
    }
}
