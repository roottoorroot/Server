using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Prototype_v1_Server
{
    public partial class Filtr : Form
    {
        //private DataGridView _dataGV;
        public string _search;

        public Filtr(DataGridView dgr)
        {
            InitializeComponent();
            initDataG();
        }

        
        private void фамилияToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            label1.Text = "Фамилия";

        }

        private void отделениеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            label1.Text = "Отделение";
        }

        private void датаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            label1.Text = "Дата";
        }

        private void состояниеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            label1.Text = "Состояние";
        }

        private void другоеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            groupBox1.Visible = true;
            label1.Text = "Ввидите ключевые слова";
        }

        private void initDataG()
        {
            dataGridView1.Columns.Add("A0", "№");
            dataGridView1.Columns.Add("A1", "Время открытия");
            dataGridView1.Columns.Add("A2", "Время Закрытия");
            dataGridView1.Columns.Add("A3", "Инженер");
            dataGridView1.Columns.Add("A4", "Отделение");
            dataGridView1.Columns.Add("A5", "Краткое описание");
            dataGridView1.Columns.Add("A6", "Полное описание");
            dataGridView1.Columns.Add("A7", "Состояние");
            dataGridView1.Columns[0].Width = 20;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 120;
            dataGridView1.Columns[3].Width = 90;
            dataGridView1.Columns[4].Width = 80;
            dataGridView1.Columns[5].Width = 150;
            dataGridView1.Columns[6].Width = 190;
            dataGridView1.Columns[7].Width = 80;
        }


        //private void sopostavlenie(DataGridView data)
        //{
        //    for (int i = 0; i < data.RowCount; i++)
        //    {
        //        for (int j = 0; j < data.ColumnCount; j++)
        //        {
        //            dataGridView1.Rows[i].Cells[j].Value = data.Rows[i].Cells[j].Value; 
        //        }
        //    }
        //}


        //private void button1_Click(object sender, EventArgs e)
        //{
        //    _search = textBox1.Text;
        //    sopostavlenie(_dataGV);
        //}
    }
}
