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

        private string selectedEmployee = "", selectedWorkDay = "", selectedWorkMonth = "", selectedWorkYear = "";
        private bool fstRunCombo2 = true, fstRunCombo3 = true;

        private List<string> arrayHolidays = new List<string>();

        private static int[] picXY = new int[2];

        Form2 employeForm;
        Form3 holidaysForm;
        Form4 macAdressesForm;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            //Size resolution = Screen.PrimaryScreen.Bounds.Size;
            Size = new Size(1220, 800);
            Top = 0;
            Left = 0;
            dataGridView1.Width = Width - 40;
            dataGridView1.Height = Convert.ToInt32(Height / 1.5);
            panel1.Width = dataGridView1.Width / 2;
            panel1.Height = comboBox1.Height * 6;
            panel2.Width = panel1.Width;
            panel2.Height = panel1.Height;
            panel2.Top = panel1.Top;
            panel2.Left = panel1.Left + panel1.Width - 1;
            button1.Top = label4.Top + label4.Height + 8;
            button1.Left = comboBox1.Left;
            button2.Top = button1.Top;
            button2.Left = button1.Left + button1.Width + 10;
            dataGridView1.Top = panel1.Top + panel1.Height - 1;
            button3.Top = dataGridView1.Top + dataGridView1.Height + 10;
            button3.Left = dataGridView1.Left;
            button4.Left = button3.Left + button3.Width + 10;
            button4.Top = button3.Top;
            button5.Left = button4.Left + button4.Width + 10;
            button5.Top = button3.Top;

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
            LoadHolidaysInArray();
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
            // Авто-ширина
            dataGridView1.AutoResizeColumns();
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            dataGridView1.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.DisableResizing;
            dataGridView1.RowHeadersVisible = false;

            comboBox2.Text = MonthToCombo[Convert.ToInt32(DateTime.Now.ToString().Substring(3, 2))];
            comboBox3.Text = DateTime.Now.ToString().Substring(6, 4);

            pictureBox1.Top = button3.Top + 4;
            pictureBox1.Left = dataGridView1.Left + dataGridView1.Width - pictureBox1.Width;
            picXY[0] = pictureBox1.Left;
            picXY[1] = pictureBox1.Top;

            label9.Top = label1.Top - 10;
            label9.Left = textBox4.Left + textBox4.Width - label9.Width;
        }
        //  ГЕТТЕР НА ПОЛОЖЕНИЕ ЛОГОТИПА
        public int[] GetPicXY()
        {
            return picXY;
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
        //  ЗАГРУЗКА В ЛОКАЛЬНЫЙ МАССИВ СПИСКА ПРАЗДНИЧНЫХ ДНЕЙ
        private void LoadHolidaysInArray()
        {
            string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
            NpgsqlConnection conn = new NpgsqlConnection(connstring);
            conn.Open();
            NpgsqlCommand comandSelect = new NpgsqlCommand("SELECT holiday FROM tblHolidays", conn);
            NpgsqlDataReader reader;
            reader = comandSelect.ExecuteReader();

            arrayHolidays.Clear();

            while (reader.Read())
            {
                arrayHolidays.Add(reader["holiday"].ToString());
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
                    textBox1.Text = startWorkTime.Substring(13, 8);
                    textBox2.Text = endWorkTime.Substring(13, 8);
                }
                else
                {
                    textBox1.Text = "-";
                    textBox2.Text = "-";
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
        //  ТАБЛИЦА ПРОСМОТРА МАК АДРЕСОВ
        private void button4_Click(object sender, EventArgs e)
        {
            macAdressesForm = new Form4();
            macAdressesForm.ShowDialog();
        }
        //  РЕДАКТИРОВАНИЕ ПРАЗДНИЧНЫХ ДНЕЙ
        private void button5_Click(object sender, EventArgs e)
        {
            holidaysForm = new Form3();
            holidaysForm.ShowDialog();

            LoadHolidaysInArray();
        }
        //  ВЫБОРКА НА МЕСЯЦ
        private void RequestData()
        {
            int daysInCurrentMonth = DateTime.DaysInMonth(Convert.ToInt32(comboBox3.Text), comboBox2.SelectedIndex + 1);
            int workDaysEmployeeInCurrentMonth = 0, holidayDaysInMonth = 0;
            bool printTheData = false;
            string day = "", today = DateTime.Today.ToShortDateString();
            if (comboBox2.SelectedIndex < 9)
                day = "01.0" + (comboBox2.SelectedIndex + 1).ToString() + "." + comboBox3.Text;
            else
                day = "01." + (comboBox2.SelectedIndex + 1).ToString() + "." + comboBox3.Text;
            string dataSt = GetTrueData(day, "start"), dataEn = GetTrueData(day, "end");
            string dataStart = "", dataEnd = "", dataForCheck = "", dataStartMonth = "", dataEndMonth = "";
            // Для прихода-ухода
            List<string> arrayTimePrihod = new List<string>(), arrayTimeUhod = new List<string>();    
            bool isFirstDay;
            string strForLastDay = "";
            // Для среднего часа
            List<int> arraySredniChas = new List<int>();
            string averHourInMonth = "";
            // Количество праздничных дней
            int holidayDays = 0;

            if ((comboBox2.SelectedIndex + 1) == Convert.ToInt32(today.Substring(3, 2)))
                daysInCurrentMonth = Convert.ToInt32(today.Substring(0, 2));

            dataGridView1.DataSource = null;
            if (IsEmployeeChecked())
            {
                string connstring = string.Format("Server={0};Port={1};User Id={2};Password={3};Database={4};", "10.0.0.99", "5432", "denver", "intGroup7", "MikrotikDb");
                NpgsqlConnection conn = new NpgsqlConnection(connstring);
                conn = new NpgsqlConnection(connstring);
                conn.Open();

                DialogResult res = MessageBox.Show("Загрузить данные по сотруднику в таблицу?", "Уведомление", MessageBoxButtons.YesNo);
                if (res == DialogResult.Yes)
                    printTheData = true;

                for (int i = 1; i <= daysInCurrentMonth; i++)
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
                        if (i == daysInCurrentMonth)
                            dataEndMonth = dataEnd;
                        dataForCheck = i.ToString() + day.Substring(2, 8);
                    }
                    if (arrayHolidays.Contains(dataForCheck))
                    {
                        holidayDays++;
                        continue;
                    }
                    if (Convert.ToDateTime(dataForCheck).DayOfWeek == DayOfWeek.Saturday || Convert.ToDateTime(dataForCheck).DayOfWeek == DayOfWeek.Sunday)
                    {
                        holidayDaysInMonth++;
                        continue;
                    }

                    NpgsqlCommand comandSelect = new NpgsqlCommand("SELECT * FROM tblData WHERE \"macDevice\" = '" + employee[comboBox1.Text] +
                                                                   "' AND \"appearTime\" > '" + dataStart + "' AND \"appearTime\" < '" + dataEnd + "' ORDER BY \"appearTime\"", conn);
                    NpgsqlDataReader reader;
                    reader = comandSelect.ExecuteReader();
                    if (reader.HasRows)
                    {
                        workDaysEmployeeInCurrentMonth++;
                        //  Приход-уход
                        int counter = 0;
                        while (reader.Read())
                        {
                            if (isFirstDay)
                            {
                                arrayTimePrihod.Add(reader["appearTime"].ToString().Substring(13, 8));
                                isFirstDay = false;
                            }
                            else
                            {
                                strForLastDay = reader["appearTime"].ToString().Substring(13, 8);
                                counter++;
                            }
                        }
                        if (counter == 0)
                            arrayTimeUhod.Add(arrayTimePrihod[arrayTimePrihod.Count - 1]);
                        else
                            arrayTimeUhod.Add(strForLastDay);
                        ///////////////////////
                    }
                    reader.Dispose();
                    //conn.Close();
                }
                //  Средний час
                double averageHoursInMonth = 0;
                for (int j = 0; j < arrayTimePrihod.Count; j++)
                    arraySredniChas.Add(ConvertTimeToSeconds(arrayTimeUhod[j]) - ConvertTimeToSeconds(arrayTimePrihod[j]));
                foreach (int val in arraySredniChas)
                    averageHoursInMonth += val;
                averageHoursInMonth /= arraySredniChas.Count;
                if (averageHoursInMonth > 0)
                    averHourInMonth = ConvertSecondsToTime(averageHoursInMonth);
                ///////////////////////
                if (printTheData)
                {
                    string sql = "SELECT * FROM tblData WHERE \"macDevice\" = '" + employee[comboBox1.Text] + "' AND \"appearTime\" > '" + dataStartMonth + "' AND \"appearTime\" < '" + dataEndMonth + "' ORDER BY \"appearTime\"";
                    NpgsqlDataAdapter da = new NpgsqlDataAdapter(sql, conn);
                    ds.Reset();
                    dt.Clear();
                    da.Fill(ds);
                    dt = ds.Tables[0];
                    dataGridView1.DataSource = dt;
                }
                if (arrayTimeUhod.Count > 0)
                {
                    textBox4.Text = (daysInCurrentMonth - holidayDaysInMonth - workDaysEmployeeInCurrentMonth - holidayDays).ToString();
                    textBox3.Text = CalculateAverTimePrihodUhod(arrayTimePrihod) + " / " + CalculateAverTimePrihodUhod(arrayTimeUhod);
                    textBox5.Text = averHourInMonth;
                }
                else
                    ClearTextBoxex("-");
                conn.Close();
            }
            else
                MessageBox.Show("Не выбран сотрудник!");
        }
        //  РАСЧЕТ ВРЕМЕНИ ПРИХОДА-УХОДА
        private string CalculateAverTimePrihodUhod(List<string> _list)
        {
            int totalSeconds = 0, sumSeconds = 0, count = 0;
            string res = "";
            double averSeconds = 0;
            // Подсчет 
            foreach(string val in _list)
            {
                if (string.Compare(val, "") != 0)
                {
                    totalSeconds = ConvertTimeToSeconds(val);
                    sumSeconds += totalSeconds;
                    count++;
                }
            }
            if (count > 0)
                averSeconds = sumSeconds / count;
            res = ConvertSecondsToTime(averSeconds);
            return res;
        }
        //  КОНВЕРТИРОВАНИЕ ВРЕМЕНИ В СЕКУНДЫ
        private int ConvertTimeToSeconds(string _time)
        {
            return(Convert.ToInt32(_time.Substring(0, 2)) * 3600 + Convert.ToInt32(_time.Substring(3, 2)) * 60 + Convert.ToInt32(_time.Substring(6, 2)));
        }
        //  КОНВЕРТИРОВАНИЕ СЕКУНД В ВРЕМЯ
        private string ConvertSecondsToTime(double _val)
        {
            string res = "";
            int h, m, s;
            double hh, mm, ss;
            hh = _val / 3600;
            h = (int)hh;
            mm = (_val - h * 3600) / 60;
            m = (int)mm;
            ss = _val - h * 3600 - m * 60;
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

        public string GetTrueData(string _d, string _f)
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
        //  ОЧИСТКА ПОЛЕЙ
        private void ClearTextBoxex(string _s)
        {
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = _s;
            textBox4.Text = _s;
            textBox5.Text = _s;
        }
        //  ВЫБОР СОТРУДНИКА
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.Compare(comboBox1.Text, selectedEmployee) != 0)
            {
                ClearTextBoxex("");
                label9.Text = comboBox1.Text;
                label9.Left = textBox4.Left + textBox4.Width - label9.Width;
                RequestData();
            }
        }
        //  ВЫБОР МЕСЯЦА
        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!fstRunCombo2)
            {
                if (string.Compare(comboBox2.Text, selectedWorkMonth) != 0)
                {
                    ClearTextBoxex("");
                    RequestData();
                }
            }
            else
                fstRunCombo2 = false;

        }
        //  ВЫБОР ГОДА
        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (!fstRunCombo3)
            {
                if (string.Compare(comboBox3.Text, selectedWorkYear) != 0)
                {
                    ClearTextBoxex("");
                    RequestData();
                }
            }
            else
                fstRunCombo3 = false;
        }
        //  СМЕНА СОТРУДНИКА
        private void comboBox1_Click(object sender, EventArgs e)
        {
            selectedEmployee = comboBox1.Text;
        }
        //  СМЕНА МЕСЯЦА
        private void comboBox2_Click(object sender, EventArgs e)
        {
            selectedWorkMonth = comboBox2.Text;
        }
        //  СМЕНА ГОДА
        private void comboBox3_Click(object sender, EventArgs e)
        {
            selectedWorkYear = comboBox3.Text;
        }
    }
}
