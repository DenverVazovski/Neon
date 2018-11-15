using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;

namespace NeonSpy
{
    public partial class Form3 : Form
    {
        Form1 frm = new Form1();

        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();

        public Form3()
        {
            InitializeComponent();
        }

        private void Form3_Load(object sender, EventArgs e)
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
            // Авто-высота
            dataGridView1.AutoResizeColumnHeadersHeight();
            dataGridView1.AutoResizeRows(DataGridViewAutoSizeRowsMode.AllCellsExceptHeaders);
            // Авто-ширина
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            LoadAllHolidays();

            int[] picXY = new int[2];
            picXY = frm.GetPicXY();
            pictureBox1.Left = picXY[0];
            pictureBox1.Top = picXY[1];
        }
        //  ЗАГРУЗКА В ТАБЛИЦУ ПРАЗДНИКОВ
        private void LoadAllHolidays()
        {
            string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
            NpgsqlConnection conn = new NpgsqlConnection(connstring);
            conn.Open();
            string sql = "SELECT * FROM tblHolidays";
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
            ds.Reset();
            da.Fill(ds);
            dt = ds.Tables[0];
            dataGridView1.DataSource = dt;
            conn.Close();
        }
        //  ЗАПИСЬ ТАБЛИЦЫ ФОРМЫ В ТАБЛИЦУ БАЗЫ ДАННЫХ
        private void WriteDataToDb()
        {
            if (dataGridView1.RowCount > 1)
            {
                string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
                NpgsqlConnection conn = new NpgsqlConnection(connstring);
                conn.Open();

                var cmd = new NpgsqlCommand("TRUNCATE tblHolidays", conn);
                cmd.ExecuteNonQuery();
                cmd.Dispose();
                NpgsqlCommand req;

                int count = 1;

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    req = new NpgsqlCommand("INSERT INTO tblHolidays VALUES(@id, @holiday)", conn);
                    req.Parameters.AddWithValue("@id", count++);
                    req.Parameters.AddWithValue("@holiday", row.Cells[1].Value.ToString());
                    req.ExecuteNonQuery();
                    req.Dispose();
                    if (count == dataGridView1.RowCount)
                        break;
                }
            }
        }
        //  ОБРАБОТЧИК КЛИКА ПРАВОЙ МЫШИ
        private void dataGridView1_CellMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridView1.ClearSelection();
            dataGridView1.Rows[e.RowIndex].Selected = true;
            dataGridView1.Focus();
            if (e.Button == MouseButtons.Right)
                contextMenuStrip1.Show(dataGridView1, new Point(Cursor.Position.X, Cursor.Position.Y));
        }
        //  ДОБАВЛЕНИЕ ПРАЗДНИЧНОГО ДНЯ
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (string.Compare(toolStripMenuItem1.Text, "Добавить") == 0)
            {
                //  РАЗРЕШЕНО РЕДАКТИРОВАНИЕ - НУЖНО НАЖАТЬ ЕЩЕ РАЗ
                dataGridView1.Rows[dataGridView1.RowCount - 2].Selected = true;
                dataGridView1.Focus();

                toolStripMenuItem1.Text = "Записать";
                dataGridView1.ReadOnly = false;
            }
            else
            {
                //  ОЧИСТКА ТАБЛИЦЫ В БАЗЕ - ЗАПИСЬ ТАБЛИЦЫ В БАЗУ
                dataGridView1.ReadOnly = true;
                toolStripMenuItem1.Text = "Добавить";
                WriteDataToDb();
                LoadAllHolidays();
            }
        }
        //  УДАЛЕНИЕ ПРАЗДНИЧНОГО ДНЯ
        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            int indexToDel = dataGridView1.SelectedCells[0].RowIndex;
            if (dataGridView1.RowCount > 1)
                dataGridView1.Rows.RemoveAt(indexToDel);
            WriteDataToDb();
            LoadAllHolidays();
        }
        //  РЕДАКТИРОВАНИЕ ПРАЗДНИЧНОГО ДНЯ
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            if (string.Compare(toolStripMenuItem3.Text, "Редактировать") == 0)
            {
                //  РАЗРЕШЕНО РЕДАКТИРОВАНИЕ - НУЖНО НАЖАТЬ ЕЩЕ РАЗ
                toolStripMenuItem3.Text = "Записать";
                dataGridView1.ReadOnly = false;
            }
            else
            {
                //  ОЧИСТКА ТАБЛИЦЫ В БАЗЕ - ЗАПИСЬ ТАБЛИЦЫ В БАЗУ
                dataGridView1.ReadOnly = true;
                toolStripMenuItem3.Text = "Редактировать";
                WriteDataToDb();
                LoadAllHolidays();
            }
        }
        //  ПРОВЕРКА ЗАПИСАНЫ ЛИ ИЗМЕНЕНИЯ В БД ПРИ ЗАКРЫТИЕ ФОРМЫ
        private void Form3_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (string.Compare(toolStripMenuItem3.Text, "Записать") == 0 || string.Compare(toolStripMenuItem1.Text, "Записать") == 0)
            {
                DialogResult res = MessageBox.Show("Запись изменений в базу данных не произведена. Вы хотите сохранить перед выходом?", "Уведомление", MessageBoxButtons.YesNo);
                if (res == DialogResult.Yes)
                    WriteDataToDb();
            }
        }
    }
}
