using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;

namespace NeonSpy
{
    public partial class Form4 : Form
    {
        Form1 frm = new Form1();

        private DataSet ds1 = new DataSet(), ds2 = new DataSet();
        private DataTable dt1 = new DataTable(), dt2 = new DataTable();

        private List<string> arrayKnownMacs = new List<string>();

        public Form4()
        {
            InitializeComponent();
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            Size = new Size(1220, 800);
            Top = 0;
            Left = 0;
            label3.Top = 12;
            label3.Left = Convert.ToInt32(Width / 3);
            label4.Top = label3.Top;
            label4.Left = Convert.ToInt32(Width / 1.225);
            dataGridView1.Width = Convert.ToInt32(Width/1.25);
            dataGridView1.Height = Convert.ToInt32(Height / 1.2);
            dataGridView2.Width = Convert.ToInt32(Width / 6.15);
            dataGridView2.Height = dataGridView1.Height;
            dataGridView1.Top = label3.Top + label3.Height + 5;
            dataGridView2.Left = dataGridView1.Left + dataGridView1.Width +5;
            dataGridView2.Top = dataGridView1.Top;

            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle
            {
                BackColor = Color.Beige,
                Font = new Font("Verdana", 10)
            };
            dataGridView2.ColumnHeadersDefaultCellStyle = dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
            dataGridView2.RowsDefaultCellStyle = dataGridView1.RowsDefaultCellStyle = columnHeaderStyle;
            // Выравнивание по центру текста
            foreach (DataGridViewColumn col in dataGridView1.Columns)
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            foreach (DataGridViewColumn col in dataGridView2.Columns)
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            // Авто-ширина
            dataGridView1.AutoResizeColumns();
            dataGridView2.AutoResizeColumns();
            dataGridView2.AutoSizeColumnsMode = dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            dataGridView2.RowHeadersWidthSizeMode = dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridView2.RowHeadersVisible = dataGridView1.RowHeadersVisible = false;

            button2.Left = dataGridView1.Left;
            button2.Top = dataGridView1.Top + dataGridView1.Height + 10;
            button1.Left = button2.Left + button2.Width + 10;
            button1.Top = button2.Top;
            label1.Left = button1.Left + button1.Width + 10;
            label1.Top = button1.Top + 4;
            dateTimePicker1.Left = label1.Left + label1.Width + 10;
            dateTimePicker1.Top = button1.Top + 2;
            label2.Left = dateTimePicker1.Left + dateTimePicker1.Width + 10;
            label2.Top = label1.Top;
            dateTimePicker2.Left = label2.Left + label2.Width + 10;
            dateTimePicker2.Top = dateTimePicker1.Top;
            
            
            int[] picXY = new int[2];
            picXY = frm.GetPicXY();
            pictureBox1.Left = picXY[0];
            pictureBox1.Top = picXY[1];

            LoadAllDevices();
            LoadKnownMacs();
        }
        //  ЗАГРУЗКА ИЗВЕСТНЫХ МАК АДРЕСОВ
        private void LoadKnownMacs()
        {
            string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};",
                                              "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
            NpgsqlConnection conn = new NpgsqlConnection(connstring);
            conn.Open();
            // Загрузка известных мак адресов
            NpgsqlCommand comandSelect = new NpgsqlCommand("SELECT macDevice FROM tblEmployee", conn);
            NpgsqlDataReader reader;
            reader = comandSelect.ExecuteReader();
            if (reader.HasRows)
                while (reader.Read())
                    arrayKnownMacs.Add(reader["macDevice"].ToString());
            reader.Dispose();
            conn.Close();
        }
        //  ЗАГРУЗКА ДАННЫХ
        private void button1_Click(object sender, EventArgs e)
        {
            string d1 = dateTimePicker1.Value.ToShortDateString(), d2 = dateTimePicker2.Value.ToShortDateString();
            int dS = Convert.ToInt32(d1.Substring(0, 2)), dE = Convert.ToInt32(d2.Substring(0, 2));

            if (dE - dS > 3)
            {
                DialogResult res = MessageBox.Show("Выбранный вами интервал выборки составляет 3 дня или более, время загрузки данных будет увеличено. Продолжить?", "Уведомление", MessageBoxButtons.YesNo);
                if (res == DialogResult.Yes)
                {
                    GetData(d1, d2);
                }
            }
            else
                GetData(d1, d2);
        }
        // ЗАГРУЗКА ДАННЫХ В ТАБЛИЦУ
        private void GetData(string _d1, string _d2)
        {
            string dataStart, dataEnd;
            dataStart = frm.GetTrueData(_d1, "start");
            dataEnd = frm.GetTrueData(_d2, "end");
            int count = 0;

            string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};",
                                              "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
            NpgsqlConnection conn = new NpgsqlConnection(connstring);
            conn.Open();
            string sql = "SELECT * FROM tblData WHERE ";
            foreach (string val in arrayKnownMacs)
            {
                if (count == arrayKnownMacs.Count - 1)
                    break;
                sql += "\"macDevice\" != '" + val + "' AND ";
                count++;
            }
            sql += "\"macDevice\" != '" + arrayKnownMacs[arrayKnownMacs.Count - 1] + "' AND \"appearTime\" > '"
                             + dataStart + "' AND \"appearTime\" < '" + dataEnd + "' ORDER BY \"macDevice\", \"appearTime\"";
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
            ds1.Reset();
            da.Fill(ds1);
            dt1 = ds1.Tables[0];
            dataGridView1.DataSource = dt1;
            conn.Close();
        }
        //  УДАЛЕНИЕ ЛИШНИХ МАКОВ
        private void button2_Click(object sender, EventArgs e)
        {
            string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
            List<string> arrayMacs = new List<string>();
            NpgsqlConnection conn = new NpgsqlConnection(connstring);
            conn.Open();
            NpgsqlCommand comandSelectAll = new NpgsqlCommand("SELECT DISTINCT \"macDevice\" FROM tblData", conn);
            NpgsqlDataReader reader;
            reader = comandSelectAll.ExecuteReader();
            //  Формирование массива мак адресов
            while (reader.Read())
            {
                if (!arrayKnownMacs.Contains(reader["macDevice"].ToString()))
                    arrayMacs.Add(reader["macDevice"].ToString());
            }
            reader.Dispose();
            comandSelectAll.Dispose();
            //  Удаление
            int count = 0, totalDelMac = 0, totalDelRec = 0;
            foreach(string val in arrayMacs)
            {
                NpgsqlCommand comandSelect = new NpgsqlCommand("SELECT count(*) FROM tblData WHERE \"macDevice\" = '" + val + "'", conn);
                NpgsqlDataReader readerDelete;
                readerDelete  = comandSelect.ExecuteReader();
                if (readerDelete != null)
                {
                    if (readerDelete.Read())
                    {
                        int.TryParse(readerDelete[0].ToString(), out count);
                    }
                }
                readerDelete.Dispose();
                comandSelect.Dispose();
                if (count < 10 && count > 0)
                {
                    NpgsqlCommand deleteCommand = new NpgsqlCommand("DELETE FROM tblData WHERE \"macDevice\" = '" + val + "'", conn);
                    deleteCommand.ExecuteNonQuery();
                    deleteCommand.Dispose();
                    totalDelMac++;
                    totalDelRec += count;
                }
            }

            //foreach (DataGridViewRow row in dataGridView2.Rows)
            for (int i = 0; i < dataGridView2.Rows.Count - 2; i++)
            {
                NpgsqlCommand deleteCommand = new NpgsqlCommand("DELETE FROM tblData WHERE \"nameDevice\" = '" + dataGridView2.Rows[i].Cells[1].Value.ToString() + "'", conn);
                deleteCommand.ExecuteNonQuery();
                deleteCommand.Dispose();
                totalDelRec++;
            }

            MessageBox.Show("Удаление завершено! Количество удаленных мак адресов - " + totalDelMac.ToString() + "\nВсего удаленно записей - " + totalDelRec);
            conn.Close();
        }
        //////////////////////////////////
        //  ЗАГРУЗКА В ТАБЛИЦУ ВСЕХ УСТРОЙСТВ
        private void LoadAllDevices()
        {
            string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
            NpgsqlConnection conn = new NpgsqlConnection(connstring);
            conn.Open();
            string sql = "SELECT * FROM tblDevices ORDER BY \"devicename\"";
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
            ds2.Reset();
            da.Fill(ds2);
            dt2 = ds2.Tables[0];
            dataGridView2.DataSource = dt2;
            conn.Close();
        }
        //  ЗАПИСЬ ТАБЛИЦЫ ФОРМЫ В ТАБЛИЦУ БАЗЫ ДАННЫХ
        private void WriteDataToDb()
        {
            if (dataGridView2.RowCount > 1)
            {
                string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
                NpgsqlConnection conn = new NpgsqlConnection(connstring);
                conn.Open();

                var cmd = new NpgsqlCommand("TRUNCATE tblDevices", conn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                NpgsqlCommand req;

                int count = 1;

                foreach (DataGridViewRow row in dataGridView2.Rows)
                {
                    req = new NpgsqlCommand("INSERT INTO tblDevices VALUES(@id, @deviceName)", conn);
                    req.Parameters.AddWithValue("@id", count++);
                    req.Parameters.AddWithValue("@deviceName", row.Cells[1].Value.ToString());
                    req.ExecuteNonQuery();
                    req.Dispose();
                    if (count == dataGridView2.RowCount)
                        break;
                }
            }
        }
        //  ДОБАВЛЕНИЕ МАКА
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (string.Compare(toolStripMenuItem1.Text, "Добавить") == 0)
            {
                //  РАЗРЕШЕНО РЕДАКТИРОВАНИЕ - НУЖНО НАЖАТЬ ЕЩЕ РАЗ
                dataGridView2.Rows[dataGridView2.RowCount - 2].Selected = true;
                dataGridView2.Focus();

                toolStripMenuItem1.Text = "Записать";
                dataGridView2.ReadOnly = false;
            }
            else
            {
                //  ОЧИСТКА ТАБЛИЦЫ В БАЗЕ - ЗАПИСЬ ТАБЛИЦЫ В БАЗУ
                dataGridView2.ReadOnly = true;
                toolStripMenuItem1.Text = "Добавить";
                WriteDataToDb();
                LoadAllDevices();
            }
        }
        //  УДАЛЕНИЕ МАКА
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            int indexToDel = dataGridView2.SelectedCells[0].RowIndex;
            if (dataGridView2.RowCount > 1)
                dataGridView2.Rows.RemoveAt(indexToDel);
            WriteDataToDb();
            LoadAllDevices();
        }
        //  РЕДАКТИРОВАНИЕ МАКА
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (string.Compare(toolStripMenuItem3.Text, "Редактировать") == 0)
            {
                //  РАЗРЕШЕНО РЕДАКТИРОВАНИЕ - НУЖНО НАЖАТЬ ЕЩЕ РАЗ
                toolStripMenuItem3.Text = "Записать";
                dataGridView2.ReadOnly = false;
            }
            else
            {
                //  ОЧИСТКА ТАБЛИЦЫ В БАЗЕ - ЗАПИСЬ ТАБЛИЦЫ В БАЗУ
                dataGridView2.ReadOnly = true;
                toolStripMenuItem3.Text = "Редактировать";
                WriteDataToDb();
                LoadAllDevices();
            }
        }
        //  ОБРАБОТЧИК КЛИКА ПРАВОЙ МЫШИ
        private void dataGridView2_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView2.ClearSelection();
            dataGridView2.Rows[e.RowIndex].Selected = true;
            dataGridView2.Focus();
            if (e.Button == MouseButtons.Right)
                contextMenuStrip1.Show(dataGridView2, PointToClient(new Point(e.X, e.Y)));
        }
        //  ЗАКРЫТИЕ ФОРМЫ
        private void Form4_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (string.Compare(toolStripMenuItem1.Text, "Записать") == 0 || string.Compare(toolStripMenuItem3.Text, "Записать") == 0)
            {
                DialogResult res = MessageBox.Show("Запись изменений в базу данных не произведена. Вы хотите сохранить перед выходом?", "Уведомление", MessageBoxButtons.YesNo);
                if (res == DialogResult.Yes)
                    WriteDataToDb();
            }
        }
    }
}
