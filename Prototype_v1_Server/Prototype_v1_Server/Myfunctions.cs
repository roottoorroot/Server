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
    class Myfunctions
    {
        public int SelectedRows(DataGridView dgv, string search, bool flag)
        {
            int _count = 0;

            for (int i = 0; i < dgv.Rows.Count; i++)
            {
                dgv.Rows[i].Selected = false;
                for (int j = 0; j < dgv.Columns.Count; j++)
                {
                    dgv.Rows[i].Cells[j].Style.BackColor = Color.White;
                }

            }

            //MessageBox.Show(Convert.ToString(_count));
            string _search = search.Replace(" ", string.Empty);
            //_search = search.Replace(".", string.Empty);
            List<DataGridViewCell> foundCell = new List<DataGridViewCell>();
            foreach (DataGridViewRow row in dgv.Rows)
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Visible && (cell.FormattedValue.ToString().Contains(search)))
                    {
                        foundCell.Add(cell);
                        _count++;
                    }
                }
            }



            if (flag == false)
            {
                for (int i = 0; i < foundCell.Count; i++)
                {
                    dgv.Rows[foundCell[i].RowIndex].Selected = true;
                }
                return _count;
            }

            else
            {




                for (int i = 0; i < dgv.Rows.Count; i++)
                {
                    for (int j = 0; j < dgv.Columns.Count; j++)
                    {
                        for (int q = 0; q < foundCell.Count; q++)
                        {
                            dgv.Rows[foundCell[q].RowIndex].Cells[j].Style.BackColor = Color.LightGreen;
                            
                        }
                    }
                }

                    return _count;
            }
            //return _count;
        }

        
        }
    

    }

