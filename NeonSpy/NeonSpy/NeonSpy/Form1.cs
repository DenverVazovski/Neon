using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Windows.Forms;
using Npgsql;

namespace NeonSpy
{
    public partial class Form1 : Form
    {

        private Dictionary<string, string> employee = new Dictionary<string, string>();
        private Dictionary<int, string> MonthToCombo = new Dictionary<int, string>();

        private DataSet ds = new DataSet();
        private DataTable dt = new DataTable();

        Form2 employeForm;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            Size resolution = Screen.PrimaryScreen.Bounds.Size;
            // if (resolution.Height == 1080)
            //    Size = new Size(1366, 768);
            //else
            //    Size = new Size(1280, 1024);
            Size = new Size(1366, 768);
            Top = 0;
            Left = 0;
            panel1.Width = this.Width - 40;
            panel2.Width = panel1.Width;
            dataGridView1.Width = panel1.Width;
            dataGridView1.Height = Convert.ToInt32(this.Height / 1.6);
            button1.Top = dataGridView1.Top + dataGridView1.Height + 10;
            button1.Left = dataGridView1.Left + dataGridView1.Width - button1.Width;
            button2.Top = button1.Top;
            button2.Left = button1.Left - button2.Width - 10;
            button3.Top = button1.Top;
            button3.Left = dataGridView1.Left;
            button4.Top = comboBox2.Top + comboBox2.Height + 10;
            button4.Left = comboBox2.Left;

            MonthToCombo.Add(01, "Январь");
            MonthToCombo.Add(02, "Февраль");
            MonthToCombo.Add(03, "Март");
            MonthToCombo.Add(04, "Апрель");
            MonthToCombo.Add(05, "Май");
            MonthToCombo.Add(06, "Июнь");
            MonthToCombo.Add(07, "Июль");
            MonthToCombo.Add(08, "Август");
            MonthToCombo.Add(09, "Сентябрь");
            MonthToCombo.Add(10, "Октябрь");
            MonthToCombo.Add(11, "Ноябрь");
            MonthToCombo.Add(12, "Декабрь");

            LoadEmployeeListInComboBox();
            // Изменение стиля оформления заголовков
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

            comboBox2.Text = MonthToCombo[Convert.ToInt32(DateTime.Now.ToString().Substring(3, 2))];
            comboBox3.Text = DateTime.Now.ToString().Substring(6, 4);
        }
        //  ПОЛУЧЕНИЕ СОТРУДНИКОВ С МАКАМИ И ЗАГРУЗКА В КОМБОБОКС
        private void LoadEmployeeListInComboBox()
        {
            string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
            NpgsqlConnection conn = new NpgsqlConnection(connstring);
            conn.Open();
            NpgsqlCommand comandSelect = new NpgsqlCommand("SELECT fullName, macDevice FROM tblEmployee", conn);
            NpgsqlDataReader reader;
            reader = comandSelect.ExecuteReader();

            employee.Clear();
            comboBox1.Items.Clear();

            while (reader.Read())
            {
                comboBox1.Items.Add(reader["fullName"].ToString());
                employee.Add(reader["fullName"].ToString(), reader["macDevice"].ToString());
            }
            reader.Dispose();
            conn.Close();

        }
        //  ПОЛНАЯ ВЫГРУЗКА
        private void button1_Click(object sender, EventArgs e)
        {
            if (IsEmployeeChecked())
            {
                string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};",
                                                  "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
                NpgsqlConnection conn = new NpgsqlConnection(connstring);
                conn.Open();
                //string sql = "SELECT count(*) FROM tblData";
                string sql = "SELECT * FROM tblData WHERE \"macDevice\" = '" + employee[comboBox1.Text] + "' ORDER BY \"appearTime\"";
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
                ds.Reset();
                da.Fill(ds);
                dt = ds.Tables[0];
                dataGridView1.DataSource = dt;
                conn.Close();
            }
            else
                MessageBox.Show("Не выбран сотрудник!");
        }
        //  ВЫГРУЗКА ОДНОГО ДНЯ
        private void button2_Click(object sender, EventArgs e)
        {
            string d = dateTimePicker1.Value.ToShortDateString();
            string dataStart, dataEnd, startWorkTime, endWorkTime;
            dataStart = GetTrueData(d, "start");
            dataEnd = GetTrueData(d, "end");

            if (IsEmployeeChecked())
            {
                string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};",
                                                  "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
                NpgsqlConnection conn = new NpgsqlConnection(connstring);
                conn.Open();
                //string sql = "SELECT count(*) FROM tblData";
                string sql = "SELECT * FROM tblData WHERE \"macDevice\" = '" + employee[comboBox1.Text] + "' AND \"appearTime\" > '"
                             + dataStart + "' AND \"appearTime\" < '" + dataEnd + "' ORDER BY \"appearTime\"";
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
                ds.Reset();
                da.Fill(ds);
                dt = ds.Tables[0];
                dataGridView1.DataSource = dt;
                conn.Close();

                if (dataGridView1.RowCount >= 2)
                { 
                    startWorkTime = dataGridView1.Rows[0].Cells[3].Value.ToString();
                    endWorkTime = dataGridView1.Rows[dataGridView1.RowCount - 2].Cells[3].Value.ToString();
                    label3.Text = "Приход на работу: " + startWorkTime.Substring(13, 8);
                    label4.Text = "Уход с работы: " + endWorkTime.Substring(13, 8);
                }
                else
                {
                    label3.Text = "Приход на работу: отсутствуют данные";
                    label4.Text = "Уход с работы: отсутствуют данные";
                }
            }
            else
                MessageBox.Show("Не выбран сотрудник!");
        }
        //  РЕДАКТИРОВАНИЕ СПИСКА СОТРУДНИКОВ
        private void button3_Click(object sender, EventArgs e)
        {
            employeForm = new Form2();
            employeForm.ShowDialog();

            LoadEmployeeListInComboBox();
        }
        //  ПОЛУЧЕНИЕ КОЛИЧЕСТВА ДНЕЙ ОТСУТСТВИЯ РАБОТНИКА
        private void button4_Click(object sender, EventArgs e)
        {
            DateTime currentDate = new DateTime();
            string day = dateTimePicker1.Value.ToShortDateString();
            string dataSt = GetTrueData(day, "start"), dataEn = GetTrueData(day, "end"), currentDay = day.Substring(0, 2);
            //string dataSt = "", dataEn = "", currentDay = day.Substring(0, 2);
            string dataStart = "", dataEnd = "", dataForCheck = "";

            int daysInCurrentMonth = DateTime.DaysInMonth(currentDate.Year, currentDate.Month);
            int workDaysEmployeeInCurrentMonth = 0;

            if (IsEmployeeChecked())
            {
                for (int i = 1; i < daysInCurrentMonth; i++)
                {
                    if (i < 10)
                    {
                        dataStart = dataSt.Replace("01", "0" + i.ToString());
                        dataEnd = dataEn.Replace("01", "0" + i.ToString());
                        dataForCheck = day.Replace("01", "0" + i.ToString());
                    }
                    else
                    {
                        dataStart = dataSt.Replace("01", i.ToString());
                        dataEnd = dataEn.Replace("01", i.ToString());
                        dataForCheck = day.Replace("01", i.ToString());
                    }
                    if ((Convert.ToDateTime(dataForCheck).DayOfWeek == DayOfWeek.Saturday) || (Convert.ToDateTime(dataForCheck).DayOfWeek == DayOfWeek.Sunday))
                        continue;

                    string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
                    NpgsqlConnection conn = new NpgsqlConnection(connstring);
                    conn.Open();
                    NpgsqlCommand comandSelect = new NpgsqlCommand("SELECT * FROM tblData WHERE \"macDevice\" = '" + employee[comboBox1.Text] +
                                                                   "' AND \"appearTime\" > '" + dataStart + "' AND \"appearTime\" < '" + dataEnd + "' ORDER BY \"appearTime\"", conn);
                    NpgsqlDataReader reader;
                    reader = comandSelect.ExecuteReader();
                    if (reader.HasRows)
                        workDaysEmployeeInCurrentMonth++;
                    reader.Dispose();
                    conn.Close();
                }
                MessageBox.Show("Дней в месяце рабочих - " + daysInCurrentMonth + "   Сотрудник отработал - " + workDaysEmployeeInCurrentMonth + "   Пропущенно дней - " + (daysInCurrentMonth - workDaysEmployeeInCurrentMonth));
            }
            else
                MessageBox.Show("Не выбран сотрудник!");
        }

        private bool IsEmployeeChecked()
        {
            if (comboBox1.Text != "")
                return true;
            return false;
        }

        private string GetTrueData(string _d, string _f)
        {
            string res = "";
            switch (_d.Substring(3, 2))
            {
                case "01":
                    res += "Jan ";
                    break;
                case "02":
                    res += "Feb ";
                    break;
                case "03":
                    res += "Mar ";
                    break;
                case "04":
                    res += "Apr ";
                    break;
                case "05":
                    res += "May ";
                    break;
                case "06":
                    res += "Jun ";
                    break;
                case "07":
                    res += "Jul ";
                    break;
                case "08":
                    res += "Aug ";
                    break;
                case "09":
                    res += "Sep ";
                    break;
                case "10":
                    res += "Oct ";
                    break;
                case "11":
                    res += "Nov ";
                    break;
                case "12":
                    res += "Dec ";
                    break;
            }

            if (string.Equals(_f, "onlyMonth"))
                return res;

            if (string.Compare(_d.Substring(0, 2), "00") > 0 && string.Compare(_d.Substring(0, 2), "10") < 0)
                res += " " + _d.Substring(1, 1);
            else
                res += _d.Substring(0, 2);

            if (string.Equals(_f, "start"))
                res += ", " + _d.Substring(6, 4) + " 00:00:00";
            else
                res += ", " + _d.Substring(6, 4) + " 23:59:59";
            return res;
        }
    }
}
