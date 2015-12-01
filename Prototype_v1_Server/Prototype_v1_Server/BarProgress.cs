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
    public partial class BarProgress : Form
    {
        public int _count;
        public BarProgress(int count)
        {
            InitializeComponent();
            _count = count;
            Progress();
        }

        public void Progress()
        {
             progressBar1.Maximum = 100;
            for (int i = 0; i < _count; i++)
            {
                progressBar1.Value = i;
                System.Threading.Thread.Sleep(100);
                
            }
        }
    }
}
