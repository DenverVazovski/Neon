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
            dataGridView1.Width = panel1.Width;
            dataGridView1.Height = Convert.ToInt32(this.Height / 1.45);
            button1.Top = dataGridView1.Top + dataGridView1.Height + 10;
            button1.Left = dataGridView1.Left + dataGridView1.Width - button1.Width;
            button2.Top = button1.Top;
            button2.Left = button1.Left - button2.Width - 10;
            button3.Top = button1.Top;
            button3.Left = dataGridView1.Left;
            button4.Top = comboBox1.Top;
            button4.Left = panel1.Left + panel1.Width - button4.Width - 30;
            button5.Top = button4.Top + button4.Height + 10;
            button5.Left = button4.Left;

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
            RequestData("otsutstvie");
           
        }
        //  СРЕДНЕЕ ВРЕМЯ ПРИХОДА И УХОДА
        private void button5_Click(object sender, EventArgs e)
        {
            //List<string> arrayTimePrihod, arrayTimeUhod;
            RequestData("prihod-uhod");
        }
        //  ВЫБОРКА НА МЕСЯЦ
        private void RequestData(string _type)
        {
            int daysInCurrentMonth = DateTime.DaysInMonth(Convert.ToInt32(comboBox3.Text), comboBox2.SelectedIndex + 1);
            int workDaysEmployeeInCurrentMonth = 0, holidayDaysInMonth = 0;
            bool printTheData = false;
            string day = "";
            if (comboBox2.SelectedIndex < 9)
                day = "01.0" + (comboBox2.SelectedIndex + 1).ToString() + "." + comboBox3.Text;
            else
                day = "01." + (comboBox2.SelectedIndex + 1).ToString() + "." + comboBox3.Text;
            string dataSt = GetTrueData(day, "start"), dataEn = GetTrueData(day, "end");
            string dataStart = "", dataEnd = "", dataForCheck = "", dataStartMonth = "", dataEndMonth = "";

            List<string> arrayTimePrihod = new List<string>(), arrayTimeUhod = new List<string>();    // Для прихода-ухода
            bool isFirstDay;
            string strForLastDay = "";

            dataGridView1.DataSource = null;
            if (IsEmployeeChecked())
            {
                DialogResult res = MessageBox.Show("Загрузить данные по сотруднику в таблицу?", "Уведомление", MessageBoxButtons.YesNo);
                if (res == DialogResult.Yes)
                    printTheData = true;

                for (int i = 1; i < daysInCurrentMonth; i++)
                {
                    isFirstDay = true;
                    if (i < 10)
                    {
                        dataStart = dataSt.Substring(0, 4) + " " + i.ToString() + dataSt.Substring(6, 15);
                        if (i == 1)
                            dataStartMonth = dataStart;
                        dataEnd = dataEn.Substring(0, 4) + " " + i.ToString() + dataEn.Substring(6, 15);
                        dataForCheck = "0" + i.ToString() + day.Substring(2, 8);
                    }
                    else
                    {
                        dataStart = dataSt.Substring(0, 4) + i.ToString() + dataSt.Substring(6, 15);
                        dataEnd = dataEn.Substring(0, 4) + i.ToString() + dataEn.Substring(6, 15);
                        if (i == daysInCurrentMonth - 1)
                            dataEndMonth = dataEnd;
                        dataForCheck = i.ToString() + day.Substring(2, 8);
                    }
                    if ((string.Compare(_type, "otsutstvie") == 0) && (Convert.ToDateTime(dataForCheck).DayOfWeek == DayOfWeek.Saturday) || (Convert.ToDateTime(dataForCheck).DayOfWeek == DayOfWeek.Sunday))
                    {
                        holidayDaysInMonth++;
                        continue;
                    }

                    string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
                    NpgsqlConnection conn = new NpgsqlConnection(connstring);
                    conn.Open();
                    NpgsqlCommand comandSelect = new NpgsqlCommand("SELECT * FROM tblData WHERE \"macDevice\" = '" + employee[comboBox1.Text] +
                                                                   "' AND \"appearTime\" > '" + dataStart + "' AND \"appearTime\" < '" + dataEnd + "' ORDER BY \"appearTime\"", conn);
                    NpgsqlDataReader reader;
                    reader = comandSelect.ExecuteReader();
                    if (reader.HasRows)
                    {
                        workDaysEmployeeInCurrentMonth++;
                        //  Приход-уход
                        if (string.Compare(_type, "prihod-uhod") == 0)
                        {
                            while (reader.Read())
                            {
                                if (isFirstDay)
                                {
                                    arrayTimePrihod.Add(reader["appearTime"].ToString().Substring(13, 8));
                                    isFirstDay = false;
                                }
                                else
                                    strForLastDay = reader["appearTime"].ToString().Substring(13, 8);
                            }
                        }
                        arrayTimeUhod.Add(strForLastDay);
                        ///////////////////////
                    }

                    reader.Dispose();
                    conn.Close();
                }
                if (printTheData)
                {
                    string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
                    NpgsqlConnection conn = new NpgsqlConnection(connstring);
                    conn.Open();
                    string sql = "SELECT * FROM tblData WHERE \"macDevice\" = '" + employee[comboBox1.Text] + "' AND \"appearTime\" > '" + dataStartMonth + "' AND \"appearTime\" < '" + dataEndMonth + "' ORDER BY \"appearTime\"";
                    NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
                    ds.Reset();
                    dt.Clear();
                    da.Fill(ds);
                    dt = ds.Tables[0];
                    dataGridView1.DataSource = dt;
                    conn.Close();
                }
                if (string.Compare(_type, "otsutstvie") == 0)
                    label8.Text = "Дней отсутствия: " + (daysInCurrentMonth - holidayDaysInMonth - workDaysEmployeeInCurrentMonth);
                else if (string.Compare(_type, "prihod-uhod") == 0)
                    label5.Text = "Среднее время (приход/уход): " + CalculateAverTimePrihodUhod(arrayTimePrihod) + " / " + CalculateAverTimePrihodUhod(arrayTimeUhod);
            }
            else
                MessageBox.Show("Не выбран сотрудник!");
        }
        //  РАСЧЕТ ВРЕМЕНИ ПРИХОДА-УХОДА
        private string CalculateAverTimePrihodUhod(List<string> _list)
        {
            int totalSeconds = 0, sumSeconds = 0, count = 0, h, m, s;
            string res = "";
            double hh, mm, ss, averSeconds;
            // Подсчет 
            foreach(string val in _list)
            {
                totalSeconds = Convert.ToInt32(val.Substring(0, 2)) * 3600 + Convert.ToInt32(val.Substring(3, 2)) * 60 + Convert.ToInt32(val.Substring(6, 2));
                sumSeconds += totalSeconds;
                count++;
            }
            averSeconds = sumSeconds / count;
            hh = averSeconds / 3600;
            h = (int)hh;
            mm = (averSeconds - h * 3600) / 60;
            m = (int)mm;
            ss = averSeconds - h * 3600 - m * 60;
            s = (int)ss;

            if (h < 10)
                res += "0" + h.ToString();
            else
                res += h.ToString();
            if (m < 10)
                res += ":0" + m.ToString();
            else
                res += ":" + m.ToString();
            if (s < 10)
                res += ":0" + s.ToString();
            else
                res += ":" + s.ToString();
            return res;
        }
        //  ПРОВЕРКА ВЫБРАН ЛИ СОТРУДНИК 
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
