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

        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();

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
            dataGridView1.Width = this.Width - 40;
            dataGridView1.Height = Convert.ToInt32(this.Height / 1.14);

            DataGridViewCellStyle columnHeaderStyle = new DataGridViewCellStyle
            {
                BackColor = Color.Beige,
                Font = new Font("Verdana", 10)
            };
            dataGridView1.ColumnHeadersDefaultCellStyle = columnHeaderStyle;
            dataGridView1.RowsDefaultCellStyle = columnHeaderStyle;
            // Выравнивание по центру текста
            foreach (DataGridViewColumn col in dataGridView1.Columns)
                col.HeaderCell.Style.Alignment = DataGridViewContentAlignment.MiddleCenter;
            // Авто-ширина
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridView1.RowHeadersVisible = false;

            button1.Left = dataGridView1.Left;
            button1.Top = dataGridView1.Top + dataGridView1.Height + 10;
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
            string d1 = dateTimePicker1.Value.ToShortDateString(), d2 = dateTimePicker1.Value.ToShortDateString();

            DialogResult res = MessageBox.Show("Выбранный вами интервал выборки 3 дня или более, время загрузки данных будет увеличено. Продолжить?", "Уведомление", MessageBoxButtons.YesNo);
            if (dateTimePicker2.Value.Date - dateTimePicker1.Value.Date > 3)
            if (res == DialogResult.Yes)
            {
                //string d1 = dateTimePicker1.Value.ToShortDateString(), d2 = dateTimePicker1.Value.ToShortDateString();
                string dataStart, dataEnd;
                dataStart = frm.GetTrueData(d1, "start");
                dataEnd = frm.GetTrueData(d2, "end");
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
                ds.Reset();
                da.Fill(ds);
                dt = ds.Tables[0];
                dataGridView1.DataSource = dt;
                conn.Close();
            }
        }
    }
}
