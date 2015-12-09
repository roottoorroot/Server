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
    public partial class Form1 : Form
    {
        
        public bool erflag = true;
        public string _patch;
        public TcpListener listener; //Listener port point
        public int count; //For DataGriWiev, rows count
        public int tec_count = 0;//Количество текущих заявок
        string resp; //Responce from klient
        Thread serverThread; //Threading for starting function: StartServer
        TcpClient client;
        

        public Form1()
        {
            InitializeComponent();
            FormingTable(1);
            //FormingTable(2);
            checkBox1.Checked = true;
            initConfig();
            DataGridVseZayavki(@"C:\log\Closelog.txt");
            
            
        }

        public void initConfig()
        {
            string patch;
            try
            {
                StreamReader reader = new StreamReader(@"C:\log\config.txt");
                patch = reader.ReadLine();
                reader.Close();
                _patch = patch;
                //label4.Text = Convert.ToString(tec_count);
            }
            catch (Exception e)
            {
                MessageBox.Show("Произошла ошибка: \r " + e.ToString());
            }
        }

        //For out from process to richTextbox
        void LogMessage(string msg)
        {
            Invoke(new Action(() =>
            {
                if (checkBox1.Checked == true) { richTextBox1.Text += msg + "\r"; }
            if (count == 0)
            {
                textBox2.Text = Convert.ToString(count);
            }
            else
            {
                textBox2.Text = Convert.ToString(count + 1);
            }
            }));
        }


        //Function for swamping filds of datagridWiev1
        void AddDataTabs(string str)
        {
            Invoke(new Action(() =>
            {
                string datatime = DateTime.Now.ToString();
                string[] list = str.Split('|');
                dataGridView1.Rows.Add();
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[0].Value = dataGridView1.Rows.Count - 1;
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[1].Value = datatime;
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[3].Value = list[0];
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[4].Value = list[1];
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[5].Value = list[2];
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[6].Value = list[3];
                dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[7].Value = "Активная";
                //Coloring allocated cells
                for (int i = 0; i < dataGridView1.ColumnCount; i++)
                {
                    dataGridView1.CurrentCell = dataGridView1.Rows[dataGridView1.Rows.Count - 2].Cells[i];
                    dataGridView1.CurrentCell.Style.BackColor = Color.LightPink;
                    dataGridView1.ClearSelection();
                }

                

            }));
            Thread infoTR = new Thread(SetBalloonTrip);
            infoTR.Start();
            //label4.Text = Convert.ToString(1);
        }

       //Function for start connection with klient in process
        public void StartServer()
        {
            //bool erflag = true;
            //int count = 0;

            //IPHostEntry ipHostInfo = Dns.Resolve(Dns.GetHostName());
            //IPAddress localhost = ipHostInfo.AddressList[0];
            IPAddress localhost = IPAddress.Parse("192.53.1.109");
            //textBox4.Text = Convert.ToString(localhost.ToString());
            listener = new TcpListener(localhost, 12345);
            listener.Start();
            LogMessage("              Ожидание входящего подключения...");
            
            while (true)
            {
                client = listener.AcceptTcpClient();
                Thread thread = new Thread(new ParameterizedThreadStart(HandleClientThread));
                thread.Start(client);
                if (erflag == false) { break; }
            }

        }

        //Obtaining response from klient
        void HandleClientThread(object obj)
        {
            string datatime = DateTime.Now.ToString();
            int new_count = 0;
            try
            {
                string received;
                TcpClient client = obj as TcpClient;
                bool done = false;
                while (!done)
                {
                    //Obtaining response from klient
                    received = ReadMessage(client);

                    //For exit from cycle(while(!done))
                    done = received.Equals("bye");

                    //Writing response from klient to RichTextBox
                    LogMessage("Incoming Message in: " + datatime + " cont: " + count);
                    label4.Text = Convert.ToString(new_count++);
                    //Writing response from klient to dataGridWiev
                    ProcessReceiv(received);
                    
                    //Definition number of incoming connection
                    count++;

                    if (done) SendResponce(client, "BYE");
                    else SendResponce(client, "OK");
                    return;
                }
                client.Close();
                
            }
            catch (Exception e)
            {
                LogMessage(e.Message);
            }
            
                
                
            
        }


        //Read klient response
        private static string ReadMessage(TcpClient client)
        {
            byte[] buffer = new byte[2*512];
            int totalRead = 0;

            do
            {
                int read = client.GetStream().Read(buffer, totalRead, buffer.Length - totalRead);
                totalRead += read;

            }
            while (client.GetStream().DataAvailable);
            return Encoding.UTF8.GetString(buffer, 0, totalRead);
        }

        //Send Responce to klient
        private static void SendResponce(TcpClient client, string message)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(message);
            client.GetStream().Write(bytes, 0, bytes.Length);
        }


        //Start server in new process
        private void button1_Click(object sender, EventArgs e)
        {
            string tnow = DateTime.Now.ToString();
            //IPHostEntry ipHostInfo = Dns.Resolve(Dns.GetHostName());
            //IPAddress localhost = ipHostInfo.AddressList[0];
            IPAddress localhost = IPAddress.Parse("192.53.1.109");
            textBox4.Text = Convert.ToString("192.53.1.109");


            textBox1.Text = "Запущен";
            richTextBox1.Text = "Статус.....\r" + "              Создание подключения.........Ок \r" + "              Сбор данных о сети...............Ок \r" + "              Запуск сервера состоялся в [" + tnow + "] \r" + "              Адрес хоста: " + localhost.ToString() + "\r";
           
            //StreamWriter logStart = new StreamWriter(_patch, true);
            //logStart.WriteLine();
            //logStart.Write(".................................Server now started in: [" + tnow + "]......................................");
            //logStart.Close();

            try
            {
                serverThread=new Thread(StartServer);
                serverThread.Name = "server";
                serverThread.Start();
                
            }
            catch (Exception ex)
            {
                richTextBox1.Text = ex.ToString();
            }
            button1.Enabled = false;
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            erflag = false;
            Thread.Sleep(0);
            //listener.Stop();//SendMailContact();
            
        }


        //Seng Message on you mail
        private void SendMailContact()
        {
            Invoke(new Action(() =>
            {
            try
            {
                SmtpClient _smtp = new SmtpClient("smtp.gmail.com", 58);
                _smtp.Credentials = new NetworkCredential("ilyafulleveline@gmail.com", "IrrEsystEblE010390");
                _smtp.EnableSsl = false;

                MailMessage message = new MailMessage();
                message.From = new MailAddress("ilyafulleveline@gmail.com");
                message.To.Add(new MailAddress("miniroot@mail.ru"));
                message.Subject = "Test";
                message.Body = "Test from mail";

                _smtp.Send(message);
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }
            }));
        }



        //Function table organization: rows & colls
        public void FormingTable(int i)
        {
            switch (i)
            {
                case 1:
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
                    break;
                case 2:
                    {
                        dataGridView2.ColumnCount = 0;
                        dataGridView2.Columns.Add("A0", "Время открытия");
                        dataGridView2.Columns.Add("A1", "Время Закрытия");
                        dataGridView2.Columns.Add("A2", "Инженер");
                        dataGridView2.Columns.Add("A3", "Отделение");
                        dataGridView2.Columns.Add("A4", "Краткое описание");
                        dataGridView2.Columns.Add("A5", "Полное описание");
                        dataGridView2.Columns.Add("A6", "...");
                        
                        dataGridView2.Columns[0].Width = 120;
                        dataGridView2.Columns[1].Width = 120;
                        dataGridView2.Columns[2].Width = 90;
                        dataGridView2.Columns[3].Width = 80;
                        dataGridView2.Columns[4].Width = 150;
                        dataGridView2.Columns[5].Width = 100;
                    }
                    break;
            }
            
        }

       
        


        //Writing response from klient to logfile
        private void LoggInFile(string patch, string received)
        {
            StreamWriter log = new StreamWriter(patch, true);
            string[] less = received.Split('|');
            string tnow = DateTime.Now.ToString();
            log.WriteLine();
            log.WriteLine(tnow + "|");
            log.WriteLine(less[0] + "|");
            log.WriteLine(less[1] + "|");
            log.WriteLine(less[2] + "|");
            log.WriteLine(less[3] + "|");
            log.Write("*");
            log.Close();
        }

        //Writing response from klient to logfile
        private void LoggInFileClose(string patch, string received)
        {
            
                StreamWriter log = new StreamWriter(patch, true);
                string[] less = received.Split('|');
                string tnow = DateTime.Now.ToString();
                log.WriteLine();
                //log.WriteLine(tnow + "|");
                //log.WriteLine(less[0] + "|");
                log.WriteLine(less[1] + "|");
                log.WriteLine(less[2] + "|");
                log.WriteLine(less[3] + "|");
                log.WriteLine(less[4] + "|");
                log.WriteLine(less[5] + "|");
                log.WriteLine(less[6] + "|");
                log.Write("*");
                log.Close();
           


            //Processing function of the incomming message
            //Removing from the incomming message
            //Adding to table dataGridWiev
            //Adding in logfile
        }
        private void ProcessReceiv(string receive)
        {

            resp = receive.Replace("\r", "").Replace("\n", "");
            AddDataTabs(resp);
            LoggInFile(_patch,resp);

        }

       

        //Processing inc message in dataGridWiev
        private void button3_Click(object sender, EventArgs e)
        {

            bool flag = false;
            int item = 0;
            string patch = @"C:\\log\\closelog.txt";
            string new_responce = "";
            string datatime = DateTime.Now.ToString();
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (dataGridView1.Rows[i].Selected)
                {
                    item = i;
                    //flag = dataGridView1.Rows[i].Selected;
                    //new_responce = SaveinExcell(1, dataGridView1);
                    dataGridView1.Rows[i].Cells[2].Value = datatime;
                    dataGridView1.Rows[i].Cells[7].Value = "Закрыта";
                    for (int z = 0; z < dataGridView1.ColumnCount; z++)
                    {

                        dataGridView1.CurrentCell = dataGridView1.Rows[i].Cells[z];
                        dataGridView1.CurrentCell.Style.BackColor = Color.LightGreen;
                        dataGridView1.ClearSelection();
                        new_responce = new_responce + dataGridView1.Rows[i].Cells[z].Value + "|";
                    }
                }
            }
            //new_responce = SaveinExcell(1, dataGridView1);
            //Здесь отдельный цикл должен быть, точнее отдельная функция
            //for (int i = 0; i < dataGridView1.RowCount; i++)
            //{
            //    if (flag)
            //    {
            //        for (int j = 0; j < dataGridView1.ColumnCount; j++)
            //        {
            //            new_responce = new_responce + dataGridView1.Rows[i].Cells[j].Value + "|";
            //        }

            //    }
            //    else
            //    {
            //        MessageBox.Show("Ни один элемент не выбран");
            //        break;
            //    }
            //}


            flag = false;
            LoggInFileClose(patch,new_responce);
            dataGridView1.Rows[item].Selected = false;
            new_responce = "";
            DataGridVseZayavki(@"C:\log\Closelog.txt");
        }


        //Отдельная функция для обработки файла в который мы запихиваем лог закрытых заявок
        public string SaveinExcell(int item, DataGridView dgv)
        {
            string responce = "";
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        responce = responce + dataGridView1.Rows[item].Cells[j].Value + "|";
                    }
            }
            return responce;
        }


        //Выход из приложения
        private void closeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dl = MessageBox.Show("Вы действительно хотите выйти из програмы?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dl == DialogResult.Yes)
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        //Показать файл лог
        private void viewLogToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
            Llog lg = new Llog(_patch);
            lg.Show(this);
            }
            catch(Exception er)
            {
                MessageBox.Show("Ошибка в Show logfile: \r" + er.ToString());
            }
        }

        //Очистить файл лог
        private void удалитьСтарыйЛогToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dl = MessageBox.Show("Вы действительно хотите очистить файл лога?", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (dl == DialogResult.Yes)
            {
                StreamWriter writeFile = new StreamWriter(_patch);
                writeFile.Write(" ");
                writeFile.Close();
                MessageBox.Show("Файл лога очищен");
            }
            
        }

        
        //охраняем данные в книге Excell
        public void saveDataInExcel(DataGridView dgr, string filename, int count)
        {
           
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            //xlWorkSheet.Columns[0,1].Width = 200;
            //xlWorkSheet.Columns[0, 2].Width = 200;
            //xlWorkSheet.Columns[0, 3].Width = 200;

            int _countRows = dgr.Rows.Count + 2; //Для прорисовки общего общего количества найденных в документе Эксель
            int i = 0;
            int j = 0;




            for (i = 0; i <= dgr.RowCount - 1; i++)
            {
                for (j = 0; j <= dgr.ColumnCount - 1; j++)
                {
                    DataGridViewCell cell = dgr[j, i];
                    xlWorkSheet.Cells[i + 2, j + 1] = cell.Value;
                    // xlWorkSheet.Range[i,j].Select();
                }
            }
            xlWorkSheet.Range["A1", "G1"].Cells.Interior.ColorIndex = 35; //Цвет заголовочной строки в экселе
            xlWorkSheet.Range["A" + _countRows, "B" + _countRows].Cells.Interior.Color = Color.LightCoral; //Цвет общего колличества входящих(Внизу)
            xlWorkSheet.Range["C2", "C" + dgr.Rows.Count].Cells.Interior.Color = Color.LightSkyBlue; //Цвет столбца инженеров

            xlWorkSheet.Cells[dgr.Rows.Count + 2, 1] = "Всего";
            xlWorkSheet.Cells[dgr.Rows.Count + 2, 2] = Convert.ToString(count);

                        xlWorkSheet.Cells[1, 1] = "Время открытия";
                        xlWorkSheet.Cells[1, 2] = "Время закрытия";
                        xlWorkSheet.Cells[1, 3] = "Инженер";
                        xlWorkSheet.Cells[1, 4] = "Отделение";
                        xlWorkSheet.Cells[1, 5] = "Краткое описание";
                        xlWorkSheet.Cells[1, 6] = "Полное описание";
                        xlWorkSheet.Cells[1, 7] = "...";


                       




                        try
                        {
                            xlWorkBook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                            xlWorkBook.Close(true, misValue, misValue);
                            xlApp.Quit();

                            releaseObject(xlWorkSheet);
                            releaseObject(xlWorkBook);
                            releaseObject(xlApp);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("Закройте файл: *.xls");
                            return;
                        }
                        //bar.Close();
        }


        //Сохраняем полученные данные в Excell
        public void saveDataInExcel(int item, string[][] dgr, string filename)
        {
            Microsoft.Office.Interop.Excel.Application xlApp;
            Microsoft.Office.Interop.Excel.Workbook xlWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet xlWorkSheet;

            object misValue = System.Reflection.Missing.Value;

            xlApp = new Microsoft.Office.Interop.Excel.Application();
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            int i = 0;
            int j = 0;

            

            for (i = 0; i < dgr.Length; i++)
            {
                for (j = 0; j < dgr[i].Length; j++)
                {
                    //string cell = dgr[i][j];
                    xlWorkSheet.Cells[i + 2, j + 1] = dgr[i][j];
                    // xlWorkSheet.Range[i,j].Select();
                }
            }
          

                xlWorkSheet.Range["A1", "G1"].Cells.Interior.ColorIndex = 35;
            //xlWorkSheet.Range["A8", "D10"].Cells.Interior.ColorIndex = 35;
                switch (item)
                {
                    case 1: 
                        {
                        xlWorkSheet.Cells[1, 1] = "Дата";
                        xlWorkSheet.Cells[1, 2] = "Инженер";
                        xlWorkSheet.Cells[1, 3] = "Отделение";
                        xlWorkSheet.Cells[1, 4] = "Краткое описание";
                        xlWorkSheet.Cells[1, 5] = "Полное описание описание";
                        }
                    break;
                    case 2:
                    {
                        xlWorkSheet.Cells[1, 1] = "Дата открытия";
                        xlWorkSheet.Cells[1, 2] = "Дата закрытия";
                        xlWorkSheet.Cells[1, 3] = "Инженер";
                        xlWorkSheet.Cells[1, 4] = "Отделение";
                        xlWorkSheet.Cells[1, 5] = "Краткое описание";
                        xlWorkSheet.Cells[1, 6] = "Полное описание";
                    }
                    break;
                }
            
            //xlWorkSheet.Cells[1, 5] = "Отделение";
            //xlWorkSheet.Cells[1, 6] = "№ Кабинета";
            //xlWorkSheet.Cells[1, 7] = "Состояние";



            try
            {
                xlWorkBook.SaveAs(filename, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);
            }
            catch (Exception e)
            {
                MessageBox.Show("Закройте файл и повторите попытку еще раз");
            }
        }
        //This hren for working saveDataIn
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

       


        private void вФайлExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult dl = MessageBox.Show("Построить отчёт", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
                    toExcellFile(1, _patch, fileName);
                    //saveDataInExcel(dataGridView1,fileName);
                    MessageBox.Show("Отчет построен", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        public void toExcellFile(int item, string patch, string fileName)
        {
            string filelog = "";
            
            try
            {
                StreamReader reader = File.OpenText(patch);
                filelog = reader.ReadToEnd();
                reader.Close();
            }
            catch (Exception e)
            {
                MessageBox.Show("Ошибка в InitText: \r" + e.ToString());
            }


            //filelog.Trim('\n');
            string[] less = filelog.Split('*');
            for (int i = 0; i < less.Length; i++)
            {
                less[i] = less[i].Replace("\r\n", string.Empty);
                
            }


            string[][] ls = new string[less.Length][];
            for (int i = 0; i < less.Length; i++)
            {
                ls[i] = less[i].Split('|');
            }
            switch (item)
            {
                case 1: saveDataInExcel(1, ls, fileName);
                    break;
                case 2: saveDataInExcel(2, ls, fileName);
                    break;
            }
            

           
        }


        private void указатьПутьКФайлуЛогToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog ofd = new FolderBrowserDialog();
            if (ofd.ShowDialog() == DialogResult.OK && ofd.SelectedPath.Length > 0)
            {
                try
                {
                    _patch = ofd.SelectedPath + @"\log.txt";

                    addFile(_patch);

                    //StreamWriter config;
                    //FileInfo file = new FileInfo(@"D:\log\config.txt");
                    //config = file.AppendText();
                    //config.WriteLine(_patch);
                    //config.Close();
                    StreamWriter config = new StreamWriter(@"C:\log\config.txt", false);
                    config.WriteLine(_patch);
                    config.Close();
                   

                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
                }
            }
        }

        public void addFile(string patch)
        {
            FileInfo fi = new FileInfo(patch);
            FileStream flstr = fi.Create();
            flstr.Close();
        }

        private void закрытыеЗаявкиxlsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            string patch = @"C:\\log\\closelog.txt";
            DialogResult dl = MessageBox.Show("Построить отчёт", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
                    toExcellFile(2, patch, fileName);
                    //saveDataInExcel(dataGridView1,fileName);
                    MessageBox.Show("Отчет построен", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void button5_Click(object sender, EventArgs e)
        {
            Filtr flt = new Filtr(dataGridView1);
            flt.Show(this);
        }
        //Функция которая выводит все заявки которые закрывались в DataGridView2
        public void DataGridVseZayavki(string patch)
        {
            string filelog = "";
            int count = 0;

            FormingTable(2);
            dataGridView2.RowCount = 1;


                try
                {
                    StreamReader reader = File.OpenText(patch);
                    filelog = reader.ReadToEnd();
                    reader.Close();
                }
                catch (Exception e)
                {
                    MessageBox.Show("Ошибка в InitText: \r" + e.ToString());
                }

            for (int i = 0; i < filelog.Length; i++)
            {
                if (filelog[i] == '*') count++;
            }



            string[] less = filelog.Split('*');
            for (int i = 0; i < less.Length; i++)
            {
                less[i] = less[i].Replace("\r\n", string.Empty);

            }

            string[][] ls = new string[less.Length][];
            for (int i = 0; i < less.Length; i++)
            {
                ls[i] = less[i].Split('|');
            }

            //dataGridView2.Columns.Add("A0", "Время открытия");
            //dataGridView2.Columns.Add("A1", "Время Закрытия");
            //dataGridView2.Columns.Add("A2", "Инженер");
            //dataGridView2.Columns.Add("A3", "Отделение");
            //dataGridView2.Columns.Add("A4", "Краткое описание");
            //dataGridView2.Columns.Add("A5", "Полное описание");
            //dataGridView2.Columns.Add("A6", "...");
            //dataGridView2.Columns[0].Width = 20;
            dataGridView2.Columns[0].Width = 120;
            dataGridView2.Columns[1].Width = 120;
            dataGridView2.Columns[2].Width = 90;
            dataGridView2.Columns[3].Width = 80;
            dataGridView2.Columns[4].Width = 150;
            dataGridView2.Columns[5].Width = 100;
            dataGridView2.Columns[6].Width = 174;
            for (int i = 0; i < count + 1;i++)
            {
                dataGridView2.Rows.Add();
            }

            
            
            for (int i = 0; i < ls.Length; i++)
            {
                for (int j = 0; j < ls[i].Length; j++)
                {
                    //string cell = dgr[i][j];
                    dataGridView2.Rows[i + 1].Cells[j].Value = ls[i][j];
                    // xlWorkSheet.Range[i,j].Select();
                }
            }

            label5.Text = Convert.ToString(dataGridView2.Rows.Count - 1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            DataGridVseZayavki(@"C:\log\Closelog.txt");
        }

        private void фильтрxlsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void button6_Click(object sender, EventArgs e)
        {
            //Filtr flt = new Filtr(dataGridView1);
            //flt.Show(this);
            button6.Enabled = false;
            groupBox3.Visible = true;
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string _search = textBox3.Text;
            int _count = 0;
            Myfunctions _functions = new Myfunctions(); //Моя библиотека;

            if (checkBox2.Checked == false)
            {
                _count = _functions.SelectedRows(dataGridView2, _search, false);
                if (_count == 0) { MessageBox.Show("Совпадений нет"); }
                else
                {
                    button9.Enabled = true;
                    label5.Text = Convert.ToString(_count);
                    label9.Text = Convert.ToString(_count);
                }
            }

            else
            {
                _count = _functions.SelectedRows(dataGridView2, _search, true);
                if (_count == 0) { MessageBox.Show("Совпадений нет"); }
                if (_count == 0) { MessageBox.Show("Совпадений нет"); }
                else
                {
                    button9.Enabled = true;
                    label5.Text = Convert.ToString(_count);
                    label9.Text = Convert.ToString(_count);
                }
                label5.Text = Convert.ToString(_count);
                label9.Text = Convert.ToString(_count);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            groupBox3.Visible = false;
            button6.Enabled = true;
        }

       
        private void button9_Click(object sender, EventArgs e)
        {
            

            DataGridView proxyDgv = new DataGridView();
            proxyDgv.Columns.Add("A0", "Время открытия");
            proxyDgv.Columns.Add("A1", "Время Закрытия");
            proxyDgv.Columns.Add("A2", "Инженер");
            proxyDgv.Columns.Add("A3", "Отделение");
            proxyDgv.Columns.Add("A4", "Краткое описание");
            proxyDgv.Columns.Add("A5", "Полное описание");
            proxyDgv.Columns.Add("A6", "...");

            int _countRows = 0;
            List<DataGridViewRow> foundRow = new List<DataGridViewRow>();
            List<DataGridViewCell> foundCell = new List<DataGridViewCell>();
            foreach (DataGridViewRow row in dataGridView2.Rows)
            {
                if (row.Selected == true)
                {
                    proxyDgv.Rows.Add();
                    foundRow.Add(row);
                    _countRows++;
                }
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Visible && (cell.Selected == true))
                    {
                        foundCell.Add(cell);
                        //_count++;
                    }
                }
            }

            int d = 0;

            for (int i = 0; i < proxyDgv.Rows.Count - 1; i++)
            {
                for (int j = 0; j < foundRow[i].Cells.Count; j++)
                {
                    proxyDgv.Rows[i].Cells[j].Value = foundRow[i].Cells[j].Value;
                }
            }
            //foreach (DataGridViewRow row in proxyDgv.Rows)
            //{
            //    foreach (DataGridViewCell cell in row.Cells)
            //    {
            //        int i = 0;
            //        cell.Value = foundCell[i].Value;
            //        i++;
            //    }
            //}
           


            DialogResult dl = MessageBox.Show("Построить отчёт", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
                    //toExcellFile(2, patch, fileName);
                    //prBar.Show(this);
                    saveDataInExcel(proxyDgv, fileName, _countRows);
                    
                    
                    
                    MessageBox.Show("Отчет построен", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void закрытьПериодToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form2 ClosePeriod = new Form2();
            ClosePeriod.Show(this);
        }

        static void SetBalloonTrip()
        {
            NotifyIcon nIcon = new NotifyIcon();
            nIcon.Icon = SystemIcons.Information;//Берем иконку из ресурса
            nIcon.Visible = true;
            nIcon.ShowBalloonTip(2000, "Заявка", "Новая заявка", ToolTipIcon.Info);
            
        }

        

        private void инфоToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Thread infoTR = new Thread(SetBalloonTrip);
            infoTR.Start();
            //SetBalloonTrip();
            
            
        }

        private void functionsToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

      
    }
}
