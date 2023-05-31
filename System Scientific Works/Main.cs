using System.Data.SqlClient;
using System;
using System.Linq;
using System.Xml.Linq;
using System.Data;
using System.Xml;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.Windows.Forms;
using System.Reflection;
using System.Windows.Markup;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.Cursor;
using System.Drawing.Printing;
using Excel = Microsoft.Office.Interop.Excel;

namespace System_Scientific_Works
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        DB db;
        Data dataF;
        Data dataD;

        private void Form1_Load(object sender, EventArgs e)
        {
            string connectionString = "Data Source=IOPC\\KURSACH;Initial Catalog=\"Scientific Works\";Integrated Security=True";

            db = new DB(connectionString);


        }

        private void displayBlock_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            db.BreakCon();
        }
        string query = "SELECT * FROM VUZ";

        private void bdChoose_SelectedIndexChanged(object sender, EventArgs e)  //Вывод данных
        {
           
            //string[] a;

            displayBlock.DataSource = null;
            SearchBy f;

            switch (bdChoose.SelectedIndex)
            {
                case 0:     //ВУЗ
                    label6.Visible = false;
                    query = "SELECT * FROM VUZ";

                    displayBlock.AllowUserToDeleteRows = true;
                    displayBlock.DataSource = db.FullTable(query).Tables[0];

                    break;
                case 1:     //Факультет
                    label6.Visible = false;
                    query = "SELECT Id, Name, Dean, Phone, Email FROM Faculty";

                    displayBlock.AllowUserToDeleteRows = true;
                    displayBlock.DataSource = db.FullTable(query).Tables[0];

                    break;
                case 2:     //Кафедра
                    label6.Visible = false;
                    query = "SELECT f.Id AS Faculty_Id, f.Name AS FacultyName, d.Id, d.Name AS Name, d.Head, d.Phone, d.Email FROM Department d " +
                        "inner join Faculty f on f.id=d.Faculty_Id " +
                        "inner join VUZ v on v.Id=f.VUZ_Id;";

                    displayBlock.AllowUserToDeleteRows = true;
                    displayBlock.DataSource = db.FullTable(query).Tables[0];

                    displayBlock.Columns["Faculty_Id"].Visible = false;
                    break;
                case 3:     //Сотрудник    
                    label6.Visible = false;
                    query = "SELECT e.Table_Number, e.FIO, e.Academic_Degree, e.Post, f.Name AS FacultyName, f.Id AS Faculty_Id, d.Name AS DepartmentName, d.Id AS Department_Id, e.Phone, e.Email FROM Employee e " +
                        "inner join Department d on d.Id = e.Department_Id " +
                        "inner join Faculty f on f.Id = d.Faculty_Id " +
                        "inner join VUZ v on v.id = f.VUZ_Id;";

                    displayBlock.AllowUserToDeleteRows = true;
                    displayBlock.DataSource = db.FullTable(query).Tables[0];

                    displayBlock.Columns["Faculty_Id"].Visible = false;
                    displayBlock.Columns["Department_Id"].Visible = false;
                    break;
                case 4:     //Количество работ сотрудников

                    query = "SELECT n.Id, e.Table_Number, e.FIO, Year, Submitted_Applications, Confirmed_Applications, Abstracts, Copyright_Certificates, Submitted_Articles,Published_Articles  FROM NumberOf n " +
                        "inner join Employee e on e.Table_Number=n.Table_Number " +
                        "inner join Department d on d.Id = e.Department_Id";

                    displayBlock.AllowUserToDeleteRows = true;
                    displayBlock.DataSource = db.FullTable(query).Tables[0];

                    displayBlock.Columns["Id"].Visible = false;

                    break;
                case 5:     //Работа
                    label6.Visible = false;
                    query = "SELECT s.Id, s.Title, m.Id AS MId, e.Table_Number AS Table_Number, e.FIO AS Author, s.UDK, s.Year, s.Type, s.NumberOfReference, s.NumberOfQuotes FROM ScientificWork s " +
                        "inner join Multiauthorship m on m.ScientificWork_Id = s.Id " +
                        "inner join Employee e on e.Table_Number = m.Table_Number;";

                    displayBlock.AllowUserToDeleteRows = true;
                    displayBlock.DataSource = db.FullTable(query).Tables[0];

                    displayBlock.Columns["Author"].ReadOnly = true;

                    displayBlock.Columns["MId"].Visible = false;
                    displayBlock.Columns["Table_Number"].Visible = false;
                    break;
                case 6:     //Запрос: для заданного факультета и года отчетности: название кафедры, число авторских свидетельств, число заявок
                    f = new SearchBy(db, bdChoose.SelectedIndex);
                    if (f.ShowDialog() == DialogResult.OK)
                    {
                        query = "SELECT a.FacultyName, a.DepartmentName, a.Year, sum(a.Copyright_Certificates) AS Copyright_Certificates, sum(a.Submitted_Applications) AS Submitted_Applications " +
                        "FROM (SELECT f.Id as fid, f.Name as FacultyName, d.Id as did, d.Name as DepartmentName, n.Year, n.Copyright_Certificates, n.Submitted_Applications " +
                        "FROM Faculty f " +
                        "inner join Department d on d.Faculty_Id=f.Id " +
                        "inner join Employee e on e.Department_Id=d.Id " +
                        "inner join NumberOf n on n.Table_Number=e.Table_Number) AS a " +
                        $"WHERE a.fid={f.Faculty} AND a.Year={f.Year} " +
                        "GROUP BY a.FacultyName, a.DepartmentName, a.Year";

                    
                        label6.Text = $"Название факультета: {f.name}\nГод: {f.Year}";
                        label6.Visible= true;

                        displayBlock.AllowUserToDeleteRows = false;
                        displayBlock.DataSource = db.FullTable(query).Tables[0];

                    }
                    break;
                case 7:     //Запрос: для всех КТН и доцентов заданного факультета: название кафедры, число отправленных статей, число опубликованных статей, число тезисов, докладов
                    f = new SearchBy(db, bdChoose.SelectedIndex);
                    if (f.ShowDialog() == DialogResult.OK)
                    {
                        query = "SELECT f.Name AS FacultyName, d.Name AS DepartmentName, e.FIO, e.Post, e.Academic_Degree, sum(n.Submitted_Articles) AS Submitted_Articles, " +
                       "sum(n.Published_Articles) AS Published_Articles, sum(n.Abstracts) AS Abstracts " +
                       "FROM Faculty f " +
                       "inner join Department d on d.Faculty_Id=f.Id " +
                       "inner join Employee e on e.Department_Id=d.Id " +
                       "inner join NumberOf n on n.Table_Number=e.Table_Number " +
                       $"WHERE e.Academic_Degree='Кандидат технических наук' AND e.Post='Доцент' AND f.Id={f.Faculty}" +
                       "GROUP BY f.Name, d.Name, e.FIO, e.Post, e.Academic_Degree";

                        displayBlock.AllowUserToDeleteRows = false;
                        displayBlock.DataSource = db.FullTable(query).Tables[0];

                        label6.Text = $"Название факультета: {f.name}";
                        label6.Visible = true;

                    }

                    break;
                case 8:     //Запрос: - для заданного факультета вывести информацию о сотрудниках, не имеющих публикаций в отчетном году: код кафедры, ФИО сотрудника, должность, ученая степень
                    f = new SearchBy(db, 6);

                    if(f.ShowDialog() == DialogResult.OK)
                    {
                        query = "SELECT d.Id AS Department_Id, d.Name AS DepartmentName, e.FIO, e.Post, e.Academic_Degree FROM Faculty f " +
                            "inner join Department d on d.Faculty_Id=f.Id " +
                            "inner join Employee e on e.Department_Id=d.Id " +
                            "inner join NumberOf n on n.Table_Number=e.Table_Number " +
                            $"WHERE n.Year = {f.Year} AND f.Id = {f.Faculty} AND n.Published_Articles=0;";

                        displayBlock.AllowUserToDeleteRows = false;
                        displayBlock.DataSource = db.FullTable(query).Tables[0];

                        label6.Text = $"Название факультета: {f.name}\nГод: {f.Year}";
                        label6.Visible = true;

                    }
                    break;

            }

            if (bdChoose.SelectedIndex<6) displayBlock.ReadOnly = false;
            else displayBlock.ReadOnly = true;

            TranslateTheTableHeader();

            displayBlock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            displayBlock.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            displayBlock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;


        }

        void TranslateTheTableHeader()
        {
            Dictionary<string, string> b = new Dictionary<string, string>();
            b.Add("Id", "Код");
            b.Add("Table_Number", "Табельный номер");
            b.Add("Name", "Название");
            b.Add("Phone", "Номер телефона");
            b.Add("Email", "Эл.почта");
            b.Add("Faculty_Id", "Код факультета");
            b.Add("Head", "Зав.Кафедры");
            b.Add("VUZ_Id", "Код ВУЗа");
            b.Add("Dean", "Декан");
            b.Add("Title", "Заголовок");
            b.Add("Rector", "Ректор");
            b.Add("Department_Id", "Код кафедры");
            b.Add("UDK", "УДК");
            b.Add("Type", "Тип работы");
            b.Add("FIO", "ФИО");
            b.Add("Post", "Должность");
            b.Add("Academic_Degree", "Науч.Степень");
            b.Add("Address", "Адрес");
            b.Add("Submitted_Applications", "Число посланных заявок");
            b.Add("Confirmed_Applications", "Число подтверждённых заявок");
            b.Add("Abstracts", "Число тезисов докладов");
            b.Add("Copyright_Certificates", "Число авторских свидетельств");
            b.Add("Submitted_Articles", "Число отправленных статей");
            b.Add("Published_Articles", "Число опубликованных статей");
            b.Add("Year", "Год");
            b.Add("NumberOfReference", "Количество ссылок");
            b.Add("NumberOfQuotes", "Количество цитирований");
            b.Add("DepartmentName", "Кафедра");
            b.Add("FacultyName", "Факультет");
            b.Add("VUZName", "ВУЗ");
            b.Add("Author", "Автор");
            b.Add("ScientificWork_Id", "Идентификатор научной работы");
            b.Add("MId", "Идентификатор многоавторства");

            for (int i = 0; i < displayBlock.Columns.Count; i++)
            {
                try
                {
                    displayBlock.Columns[i].HeaderText = b[displayBlock.Columns[i].Name];
                }
                catch { }
                if (displayBlock.Columns[i].Name == "Id"
                    || displayBlock.Columns[i].Name == "Department_Id"
                    || displayBlock.Columns[i].Name == "DepartmentName"
                    || displayBlock.Columns[i].Name == "Faculty_Id"
                    || displayBlock.Columns[i].Name == "FacultyName"
                    || displayBlock.Columns[i].Name == "VUZName"
                    || displayBlock.Columns[i].Name == "Table_Number"
                    || displayBlock.Columns[i].Name == "Post")
                {
                    displayBlock.Columns[i].ReadOnly = true;
                }
            }

        }

        private void addINTO_SelectedIndexChanged(object sender, EventArgs e) //Выбор объекта добавления
        {
            textBox1.Visible = true;
            textBox2.Visible = true;
            textBox3.Visible = true;
            label1.Visible = true;
            label2.Visible = true;
            label3.Visible = true;
            label4.Visible = true;
            button1.Visible = true;

            checkBox1.Visible = false;
            textBox9.Visible = false;
            checkBox1.Checked = false;

            textBox5.Location = new Point(comboBox1.Location.X, textBox5.Location.Y);
            textBox4.Location = new Point(comboBox1.Location.X, textBox4.Location.Y);
            textBox3.Location = new Point(comboBox1.Location.X, textBox3.Location.Y);
            textBox2.Location = new Point(comboBox1.Location.X, textBox2.Location.Y);
            textBox1.Location = new Point(comboBox1.Location.X, textBox1.Location.Y);
            typeBox.Location = new Point(typeBox2.Location.X, typeBox.Location.Y);

            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox9.Clear();

            switch (addINTO.SelectedIndex) 
            {
                case 0:     //Факультет
                    HideShow(addINTO.SelectedIndex);
                    label4.Text = "Декан";

                    break;

                case 1:     //Кафедра
                    HideShow(addINTO.SelectedIndex);
                    label4.Text = "Зав.кафедры";
                    label5.Text = "Факультет";

                    comboBox2.Items.Clear();
                    dataF = db.GetNameId("Faculty");
                    foreach (var t in dataF.name) comboBox2.Items.Add(t);

                    break;

                case 2:     //Сотрудник
                    HideShow(addINTO.SelectedIndex);

                    label1.Text = "ФИО";
                    label4.Text = "Факультет";
                    label5.Text = "Кафедра";
                    label9.Text = "Учёная степень";

                    typeBox.DropDownStyle = ComboBoxStyle.DropDown;
                    typeBox.Items.Clear();
                    typeBox.Items.Add("Кандидат наук");
                    typeBox.Items.Add("Кандидат технических наук");
                    typeBox.Items.Add("Доктор наук");
                    typeBox.Items.Add("Доктор технических наук");

                    comboBox1.Items.Clear();
                    dataF = db.GetNameId("Faculty");
                    foreach (var t in dataF.name) comboBox1.Items.Add(t);

                    comboBox2.Items.Clear();
                    dataD = db.GetNameId("Department");
                    foreach (var t in dataD.name) comboBox2.Items.Add(t);

                    break;

                case 3:     //Научный труд
                    HideShow(addINTO.SelectedIndex);
                    checkBox1.Visible = true;
                    textBox9.Visible = true;

                    label1.Text = "Заголовок";
                    label2.Text = "УДК";
                    label3.Text = "Год";
                    label4.Text = "Количество ссылок";
                    label5.Text = "Количество цитирований";
                    label9.Text = "Тип";

                    textBox5.Location = new Point(textBox5.Location.X + 90, textBox5.Location.Y);
                    textBox4.Location = new Point(textBox4.Location.X + 90, textBox4.Location.Y);
                    textBox3.Location = new Point(textBox3.Location.X - 60, textBox3.Location.Y);
                    textBox2.Location = new Point(textBox2.Location.X - 60, textBox2.Location.Y);
                    textBox1.Location = new Point(textBox1.Location.X - 25, textBox1.Location.Y);
                    typeBox.Location = new Point(typeBox.Location.X - 75, typeBox.Location.Y);

                    typeBox.DropDownStyle = ComboBoxStyle.DropDownList;
                    typeBox.Items.Clear();
                    typeBox.Items.Add("Статья");
                    typeBox.Items.Add("Реферат");
                    typeBox.Items.Add("Диссертация");
                    typeBox.Items.Add("Эссе");
                    typeBox.Items.Add("Отчёт о НИР");
                    typeBox.Items.Add("Докторская");
                    typeBox.Items.Add("Книга");
                    typeBox.Items.Add("Учебное издание");
                    typeBox.Items.Add("Руководство пользователя");
                    typeBox.Items.Add("Авторское свидетельство");

                    break;

            }
        }

        private void HideShow(int i)    //Проявление и скрытие полей вкладки добавления
        {
            if (i < 2) label1.Text = "Название";

            if (i == 2)
            {
                typeBox2.Visible = true;
                label10.Visible = true;
                textBox4.Visible = false;
                comboBox1.Visible = true;
            }
            else
            {
                typeBox2.Visible = false;
                label10.Visible = false;
                textBox4.Visible = true;
                comboBox1.Visible = false;
            }

            if (i != 3)
            {
                label2.Text = "Тел.";
                label3.Text = "Эл.почта";
            }

            if (i != 0)
            {
                label5.Visible = true;
                if (i == 2 || i == 3) label9.Visible = true;
                else label9.Visible = false;
                if (i != 3)
                {
                    comboBox2.Visible = true;
                    textBox5.Visible = false;
                    label12.Visible = false;
                    textBox6.Visible = false;
                    label13.Visible = false;
                }
                else
                {
                    comboBox2.Visible = false;
                    textBox5.Visible = true;
                    label12.Visible=true; 
                    textBox6.Visible=true; 
                    label13.Visible=true;
                }
                if (i > 1) typeBox.Visible = true;
                else typeBox.Visible = false;
            }
            else
            {
                textBox5.Visible = false;
                label5.Visible = false;
                label9.Visible = false;
                comboBox2.Visible = false;
                typeBox.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)  //Добавление записей
        {
            string q;
            switch (addINTO.SelectedIndex)
            {
                case 0:     //Факультет
                    if (string.IsNullOrEmpty(textBox1.Text))
                    {
                        MessageBox.Show("Пожалуйста, убедитесь, что поле <Название> заполненно верно", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    try { Convert.ToInt64(textBox2.Text); }
                    catch { MessageBox.Show("Пожалуйста, убедитесь, что поле <Тел.> содержит в себе только цифры", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }


                    q = $"'{textBox1.Text}', ";
                    if(string.IsNullOrEmpty(textBox4.Text)) q += "NULL";
                    else q += $"'{textBox4.Text}'";
                    q += ", ";
                    if (string.IsNullOrEmpty(textBox2.Text)) q += "NULL";
                    else q += textBox2.Text;
                    q += ", ";
                    if (string.IsNullOrEmpty(textBox3.Text)) q += "NULL";
                    else q += $"'{textBox3.Text}'";

                    db.DoQuery($"INSERT INTO Faculty (Name, Dean, Phone, Email, VUZ_Id) VALUES({q}, 1)");

                    break;

                case 1:     //Кафедра
                    if (string.IsNullOrEmpty(textBox1.Text))
                    {
                        MessageBox.Show("Пожалуйста, убедитесь, что поле <Название> заполненно верно", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if(comboBox2.SelectedIndex == -1)
                    {
                        MessageBox.Show("Пожалуйста, убедитесь, что был выбран факультет кафедры", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    
                    try { Convert.ToInt64(textBox2.Text); }
                    catch { MessageBox.Show("Пожалуйста, убедитесь, что поле <Тел.> содержит в себе только цифры", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                    q = $"'{textBox1.Text}', ";
                    if (string.IsNullOrEmpty(textBox4.Text)) q += "NULL";
                    else q += $"'{textBox4.Text}'";
                    q += ", ";
                    if (string.IsNullOrEmpty(textBox2.Text)) q += "NULL";
                    else q += textBox2.Text;
                    q += ", ";
                    if (string.IsNullOrEmpty(textBox3.Text)) q += "NULL";
                    else q += $"'{textBox3.Text}'";

                    db.DoQuery($"INSERT INTO Department (Name, Head, Phone, Email, Faculty_Id) VALUES({q}, {dataF.id[comboBox2.SelectedIndex]})");

                    break;

                case 2: //Сотрудник
                    if (string.IsNullOrEmpty(textBox1.Text))
                    {
                        MessageBox.Show("Пожалуйста, убедитесь, что поле <ФИО> заполненно верно", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }
                    if(comboBox2.SelectedIndex == -1)
                    {
                        MessageBox.Show("Пожалуйста, убедитесь, что была выбрана кафедра сотрудника", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        return;
                    }

                    try { Convert.ToInt64(textBox2.Text); }
                    catch { MessageBox.Show("Пожалуйста, убедитесь, что поле <Тел.> содержит в себе только цифры", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }

                    q = $"'{textBox1.Text}', ";
                    if (string.IsNullOrEmpty(typeBox.Text)) q += "NULL";
                    else q += $"'{typeBox.Text}'";
                    q += ", ";
                    if (string.IsNullOrEmpty(typeBox2.Text)) q += "NULL";
                    else q += $"'{typeBox2.Text}'";
                    q += ", ";
                    if (string.IsNullOrEmpty(textBox2.Text)) q += "NULL";
                    else q += textBox2.Text;
                    q += ", ";
                    if (string.IsNullOrEmpty(textBox3.Text)) q += "NULL";
                    else q += $"'{textBox3.Text}'";

                    db.DoQuery($"INSERT INTO Employee (FIO, Academic_Degree, Post, Phone, Email, Department_Id) VALUES({q}, {dataD.id[comboBox2.SelectedIndex]})");
                    
                    //Добавление связанного объекта Число
                    db.DoQuery($"INSERT INTO NumberOf (Table_Number, Year) VALUES({db.GetLastId("Employee","Table_Number")}, {DateTime.Now.Year})");

                    break;

                case 3: //Научный труд
                    if(checkBox1.Checked) 
                    { 
                        if(string.IsNullOrEmpty(textBox9.Text))
                        {
                            MessageBox.Show("Пожалуйста, убедитесь, что поле <Идентификатор работы> заполненно верно", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        if (string.IsNullOrEmpty(textBox2.Text))
                        {
                            MessageBox.Show("Пожалуйста, убедитесь, что все поле <Табельный номер> заполненно верно", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        try
                        {
                            int a = Convert.ToInt32(textBox2.Text);
                            int b = Convert.ToInt32(textBox9.Text);

                            //Добавить проверку на наличие работы и сотрудника

                            db.DoQuery($"INSERT INTO Multiauthorship (Table_Number, ScientificWork_Id) VALUES({textBox2.Text}, {textBox9.Text})");
                        }
                        catch (FormatException) 
                        {
                            MessageBox.Show("Поля <Табельный номер> и <Идентификатор работы> могут принимать только числовые значения", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return; 
                        }
                        catch(Exception ex) 
                        {
                            MessageBox.Show("Возможно, введённые вами идентификатор работы или табельный номер сотрудника отсуствуют в базе данных\n\n"+ex.Message, "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                    }
                    else
                    {
                        if (string.IsNullOrEmpty(textBox1.Text))
                        {
                            MessageBox.Show("Пожалуйста, убедитесь, что поле <Заголовок> заполненно верно", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        if (string.IsNullOrEmpty(textBox3.Text))
                            textBox3.Text = DateTime.Now.Year.ToString();

                        if (string.IsNullOrEmpty(textBox6.Text))
                        {
                            MessageBox.Show("Пожалуйста, убедитесь, что поле <Табельный номер сотрудника> заполненно верно", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }

                        q = $"'{textBox1.Text}', ";
                        if (string.IsNullOrEmpty(textBox2.Text)) q += "NULL";
                        else q += $"'{textBox2.Text}'";
                        q += $", {textBox3.Text}, ";
                        if (string.IsNullOrEmpty(textBox4.Text)) q += "0";
                        else q += textBox4.Text;
                        q += ", ";
                        if (string.IsNullOrEmpty(textBox5.Text)) q += "0";
                        else q += textBox5.Text;
                        q += ", ";
                        if (string.IsNullOrEmpty(typeBox.Text)) q += "NULL";
                        else q += $"'{typeBox.Text}'";

                        string[] authors = textBox6.Text.Trim().Split(',');

                        try
                        {
                            int a = Convert.ToInt32(textBox3.Text);
                            int[] table_nums = new int[authors.Length];
                            for(int i =0; i < table_nums.Length; i++)
                                table_nums[i] = Convert.ToInt32(authors[i]);
                            int c, d;
                            if(!string.IsNullOrEmpty(textBox4.Text)) c = Convert.ToInt32(textBox4.Text);
                            if (!string.IsNullOrEmpty(textBox5.Text)) d = Convert.ToInt32(textBox5.Text);

                            //Добавить проверку на наличие работы и сотрудника

                            db.DoQuery($"INSERT INTO ScientificWork (Title, UDK, Year, NumberOfReference, NumberOfQuotes, Type) VALUES({q});");

                            int id = db.GetLastId("ScientificWork", "Id");

                            foreach(var t in table_nums) db.DoQuery($"INSERT INTO Multiauthorship (Table_Number, ScientificWork_Id) VALUES({t}, {id})");

                        }
                        catch (FormatException)
                        {
                            MessageBox.Show("Поля <Год>, <Табельный номер>, <Количество цитирований> и <Количество ссылок> могут принимать только числовые значения", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        catch (Exception ex) 
                        {
                            MessageBox.Show("Убедитесь, что табельный номер сотрудника указан верно\n\n"+ex.Message, "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            return;
                        }
                        
                    }

                    break;

            }
            MessageBox.Show("Запись была успешно добавлена в базу данных", "Успех", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e) //Добавление автора к работе
        {
            if (checkBox1.Checked)
            {
                textBox9.Enabled = true;

                label1.Visible = false;
                textBox1.Visible = false;
                label3.Visible = false;
                textBox3.Visible = false;
                label4.Visible = false;
                textBox4.Visible = false;
                label5.Visible = false;
                textBox5.Visible = false;
                label9.Visible = false;
                typeBox.Visible = false;
                label12.Visible = false;
                label13.Visible = false;
                textBox6.Visible = false;

                label2.Text = "Табельный номер";
                textBox2.Location = new Point(textBox2.Location.X + 100, textBox2.Location.Y);
            }
            else
            {
                textBox2.Location = new Point(textBox1.Location.X, textBox2.Location.Y);
                textBox9.Enabled = false;

                addINTO_SelectedIndexChanged(null, null);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)     //Поиск кафедры относительно выбранного факультета
        {
            comboBox2.Items.Clear();
            dataD = db.GetNameId("Department", dataF.id[comboBox1.SelectedIndex]);
            foreach (var t in dataD.name) comboBox2.Items.Add(t);
        }


        string cell_content;
        string table_name;
        string column_name;
        string where;
        string new_name;
        int row_numb;
        bool updated = false;

        private void displayBlock_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e) // Вывод вариантов подстановки значений в виде контекстного меню
        {
            cell_content = displayBlock[e.ColumnIndex, e.RowIndex].Value.ToString();

            if (displayBlock.Columns[e.ColumnIndex].Name == "Academic_Degree")
            {
                ContextMenuStrip lkm_menu = new ContextMenuStrip();
                lkm_menu.Items.Add("Кандидат наук").Name = "Academic_Degree";
                lkm_menu.Items.Add("Кандидат технических наук").Name = "Academic_Degree";
                lkm_menu.Items.Add("Доктор наук").Name = "Academic_Degree";
                lkm_menu.Items.Add("Доктор технических наук").Name = "Academic_Degree";
                lkm_menu.Show(Cursor.Position.X, Cursor.Position.Y);

                column_name = displayBlock.Columns[e.ColumnIndex].Name;
                table_name = "Employee";
                for (int i = 0; i < displayBlock.Columns.Count; i++)
                    if (displayBlock.Columns[i].Name == "Table_Number")
                        where = "Table_Number=" + displayBlock[i, e.RowIndex].Value;

                row_numb = e.RowIndex;
                lkm_menu.ItemClicked += new ToolStripItemClickedEventHandler(lkm_menu_ItemClicked);
            }

            if (displayBlock.Columns[e.ColumnIndex].Name == "Type")
            {
                ContextMenuStrip lkm_menu = new ContextMenuStrip();
                lkm_menu.Items.Add("Статья").Name = "Type";
                lkm_menu.Items.Add("Реферат").Name = "Type";
                lkm_menu.Items.Add("Диссертация").Name = "Type";
                lkm_menu.Items.Add("Эссе").Name = "Type";
                lkm_menu.Items.Add("Отчёт о НИР").Name = "Type";
                lkm_menu.Items.Add("Докторская").Name = "Type";
                lkm_menu.Items.Add("Книга").Name = "Type";
                lkm_menu.Items.Add("Учебное издание").Name = "Type";
                lkm_menu.Items.Add("Руководство пользователя").Name = "Type";
                lkm_menu.Items.Add("Патент").Name = "Type";

                lkm_menu.Show(Cursor.Position.X, Cursor.Position.Y);

                column_name = displayBlock.Columns[e.ColumnIndex].Name;
                table_name = "ScientificWork";
                for (int i = 0; i < displayBlock.Columns.Count; i++)
                    if (displayBlock.Columns[i].Name == "Id")
                        where = "Id=" + displayBlock[i, e.RowIndex].Value;

                row_numb = e.RowIndex;
                lkm_menu.ItemClicked += new ToolStripItemClickedEventHandler(lkm_menu_ItemClicked);
            }
        }

        

        private void lkm_menu_ItemClicked(object? sender, ToolStripItemClickedEventArgs e) // Обработчик пунктов контекстного меню
        {
            //throw new NotImplementedException();
            if (MessageBox.Show("Вы действительно хотите обновить запись базы данных", "Обновление", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                string what;

                if (bdChoose.SelectedIndex == 2 && column_name == "FacultyName")
                {
                    what = $"Faculty_Id={dataF.id[Convert.ToInt32(e.ClickedItem.Tag)]}";
                    db.Update(table_name, what, where);
                    displayBlock[column_name, row_numb].Value = dataF.name[Convert.ToInt32(e.ClickedItem.Tag)];
                }
                if(bdChoose.SelectedIndex == 3)
                {

                }

                if (bdChoose.SelectedIndex < 5 && column_name != "FacultyName" && column_name != "DepartmentName")
                {
                    new_name = e.ClickedItem.Text;
                    what = column_name + $" = '{new_name}'";
                    db.Update(table_name, what, where);
                    MessageBox.Show("База данных обновлена", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    updated = true;
                    displayBlock[column_name, row_numb].Value = new_name;
                }
               
            }
        }

        private void displayBlock_CellEndEdit(object sender, DataGridViewCellEventArgs e) // Выход из редактирования и обновление сведений
        {
            if (displayBlock[e.ColumnIndex, e.RowIndex].Value.ToString() != cell_content && !updated)
            {
                string what = "";
                column_name = displayBlock.Columns[e.ColumnIndex].Name;

                if (column_name == "Phone")
                {
                    if(displayBlock[e.ColumnIndex, e.RowIndex].Value.ToString().Trim().Length >= 10)
                    {
                        try
                        {
                            new_name = displayBlock[e.ColumnIndex, e.RowIndex].Value.ToString().Trim();
                            Convert.ToInt64(new_name);
                            what = column_name + $" = {new_name}";
                        }
                        catch (FormatException)
                        {
                            displayBlock[e.ColumnIndex, e.RowIndex].Value = cell_content;
                            MessageBox.Show("Номер должен содержать только цифры", "Ошибка ввода", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }
                    else
                    {
                        displayBlock[e.ColumnIndex, e.RowIndex].Value = cell_content;
                        MessageBox.Show("Убедитесь, что номер телефона указан целиком", "Недопустимая длина номера телефона", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                if (MessageBox.Show("Вы точно уверены, что хотите изменить содержимое этой ячейки?", "Внимание!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    new_name = displayBlock[e.ColumnIndex, e.RowIndex].Value.ToString().Trim();  
                    
                    switch (bdChoose.SelectedIndex)
                    {
                        case 0:
                            table_name = "VUZ";

                            if(column_name != "Phone")
                                what = column_name + $" = '{new_name}'";
                            try
                            {
                                db.Update(table_name, what, "Id=1");
                            }
                            catch(Exception ex) 
                            {
                                MessageBox.Show("Убедитесь, что вы ввели число\n\n" + ex.Message, "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return; 
                            }
                            
                            break;

                        case 1:
                            table_name = "Faculty";

                            if (column_name != "Phone")
                                what = column_name + $" = '{new_name}'";

                            //for (int i = 0; i < displayBlock.Columns.Count; i++)
                            for (int i = 0; i < displayBlock.Columns.Count; i++)
                                if (displayBlock.Columns[i].Name == "Id")
                                    where = "Id=" + displayBlock[i, e.RowIndex].Value;                        

                            try
                            {
                                db.Update(table_name, what, where);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Убедитесь, что вы ввели число\n\n" + ex.Message, "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            break;

                        case 2:
                            table_name = "Department";

                            if (column_name != "Phone")
                                what = column_name + $" = '{new_name}'";

                            for (int i = 0; i < displayBlock.Columns.Count; i++)
                                if (displayBlock.Columns[i].Name == "Id")
                                    where = "Id=" + displayBlock[i, e.RowIndex].Value;

                            try
                            {
                                db.Update(table_name, what, where);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Убедитесь, что вы ввели число\n\n" + ex.Message, "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            break;

                        case 3:
                            table_name = "Employee";

                            if (column_name != "Phone")
                                what = column_name + $" = '{new_name}'";

                            for (int i = 0; i < displayBlock.Columns.Count; i++)
                                if (displayBlock.Columns[i].Name == "Id")
                                    where = "Id=" + displayBlock[i, e.RowIndex].Value;

                            try
                            {
                                db.Update(table_name, what, where);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Убедитесь, что вы ввели число\n\n" + ex.Message, "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            break;

                        case 4:
                            table_name = "NumberOf";
                            if (column_name != "Phone")
                                what = column_name + $" = '{new_name}'";

                            for (int i = 0; i < displayBlock.Columns.Count; i++)
                                if (displayBlock.Columns[i].Name == "Id")
                                    where = "Id=" + displayBlock[i, e.RowIndex].Value;
                            
                            db.Update(table_name, what, where);
                            break;

                        case 5:
                            table_name = "ScientificWork";

                            if (column_name != "Phone")
                                what = column_name + $" = '{new_name}'";

                            for (int i = 0; i < displayBlock.Columns.Count; i++)
                                if (displayBlock.Columns[i].Name == "Id")
                                    where = "Id=" + displayBlock[i, e.RowIndex].Value;

                            try
                            {
                                db.Update(table_name, what, where);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show("Убедитесь, что вы ввели число\n\n" + ex.Message, "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                return;
                            }

                            break;

                    }

                    for (int i = 0; i < displayBlock.Columns.Count; i++)
                        if (displayBlock.Columns[i].Name == "Table_Number")
                            where = "Table_Number=" + displayBlock[i, e.RowIndex].Value;

                    //db.Update(table_name, what, where);
                    MessageBox.Show("Данные были успешно обновлены!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    displayBlock[e.ColumnIndex, e.RowIndex].Value = new_name;
                }
                else displayBlock[e.ColumnIndex, e.RowIndex].Value = cell_content;

            }

            updated = false;
        }

        

        private void displayBlock_CellDoubleClick(object sender, DataGridViewCellEventArgs e)       //Изменение данных через контектное меню
        {
            if(e.ColumnIndex >=0 && e.RowIndex>=0 && bdChoose.SelectedIndex < 6)
            {
                if (displayBlock[e.ColumnIndex, e.RowIndex].ReadOnly == true)
                {
                    if (displayBlock.Columns[e.ColumnIndex].Name == "Post")
                    {
                        ContextMenuStrip lkm_menu = new ContextMenuStrip();
                        lkm_menu.Items.Add("Доцент").Name = "Post";
                        lkm_menu.Items.Add("Профессор").Name = "Post";
                        lkm_menu.Show(Cursor.Position.X, Cursor.Position.Y);

                        column_name = displayBlock.Columns[e.ColumnIndex].Name;
                        table_name = "Employee";
                        for (int i = 0; i < displayBlock.Columns.Count; i++)
                            if (displayBlock.Columns[i].Name == "Table_Number")
                                where = "Table_Number=" + displayBlock[i, e.RowIndex].Value;

                        row_numb = e.RowIndex;

                        lkm_menu.ItemClicked += new ToolStripItemClickedEventHandler(lkm_menu_ItemClicked);

                    }

                    if (displayBlock.Columns[e.ColumnIndex].Name == "FacultyName" && bdChoose.SelectedIndex == 2)
                    {
                        ContextMenuStrip lkm_menu = new ContextMenuStrip();
                        dataF = db.GetNameId("Faculty");
                        for (int i = 0; i < dataF.id.Count; i++)
                        {
                            lkm_menu.Items.Add(dataF.name[i]).Tag = i.ToString();
                        }

                        lkm_menu.Show(Cursor.Position.X, Cursor.Position.Y);

                        column_name = displayBlock.Columns[e.ColumnIndex].Name;
                        table_name = "Department";
                        for (int i = 0; i < displayBlock.Columns.Count; i++)
                            if (displayBlock.Columns[i].Name == "Id")
                                where = "Id=" + displayBlock[i, e.RowIndex].Value;

                        row_numb = e.RowIndex;

                        lkm_menu.ItemClicked += new ToolStripItemClickedEventHandler(lkm_menu_ItemClicked);
                    }

                    if ((displayBlock.Columns[e.ColumnIndex].Name == "DepartmentName" || displayBlock.Columns[e.ColumnIndex].Name == "FacultyName") && bdChoose.SelectedIndex == 3) 
                    {
                        ChangeTo chg = new ChangeTo(db, "SELECT d.Faculty_Id, f.Name AS 'Название факультета', d.Id AS 'Код кафедры', d.Name AS 'Название кафедры' FROM Department d inner join Faculty f on f.Id=d.Faculty_Id;");
                        if (chg.ShowDialog() == DialogResult.OK)
                        {
                            column_name = displayBlock.Columns[e.ColumnIndex].Name;
                            table_name = "Employee";

                            db.Update(table_name, $"Department_Id={chg.newId}", $"Table_Number={displayBlock["Table_Number", e.RowIndex].Value}");
                            displayBlock["DepartmentName", e.RowIndex].Value = chg.newName;
                            displayBlock["FacultyName", e.RowIndex].Value = chg.newName2;
                            MessageBox.Show("Данные были успешно обновлены!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        }
                    }

                    if(displayBlock.Columns[e.ColumnIndex].Name == "Author")
                    {
                        ChangeTo chg = new ChangeTo(db, "Employee", "Table_Number, FIO");
                        if (chg.ShowDialog() == DialogResult.OK)
                        {
                            db.Update("Multiauthorship", $"Table_Number={chg.newId}", $"Id={displayBlock["MId", e.RowIndex].Value}");

                            displayBlock[e.ColumnIndex, e.RowIndex].Value = chg.newName;
                            MessageBox.Show("Данные были успешно обновлены!", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                             
                        }
                    }

                }
            }
        }

        private void displayBlock_RowsRemoved(object sender, DataGridViewRowsRemovedEventArgs e)
        {

           
        }

        private void displayBlock_UserDeletingRow(object sender, DataGridViewRowCancelEventArgs e)      //Удаление записи с БД
        {
            switch (bdChoose.SelectedIndex)
            {
                case 1:
                    table_name = "Faculty";

                    if (MessageBox.Show("Вы уверены, что хотите удалить данную запись. В случае наличия связанных элементов <Кафедра>, <Сотрудник>, они будут также удалены!", "Внимание!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        db.Delete(table_name, $"Id={displayBlock["Id", e.Row.Index].Value}");
                        MessageBox.Show("Удаление прошло успешно", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else e.Cancel = true;

                    break;

                case 2:
                    table_name = "Department";

                    if (MessageBox.Show("Вы уверены, что хотите удалить данную запись. В случае наличия связанных элементов <Сотрудник>, они будут также удалены!", "Внимание!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        db.Delete(table_name, $"Id={displayBlock["Id", e.Row.Index].Value}");
                        MessageBox.Show("Удаление прошло успешно", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else e.Cancel = true;
                    break;

                case 3:
                    table_name = "Employee";

                    if (MessageBox.Show("Вы уверены, что хотите удалить данную запись. В случае наличия связанных элементов <Научный труд>, они останутся без автора!", "Внимание!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        db.Delete(table_name, $"Table_Number={displayBlock["Table_Number", e.Row.Index].Value}");
                        db.Delete("MultiAuthorship", $"Table_Number={displayBlock["Table_Number", e.Row.Index].Value}");
                        MessageBox.Show("Удаление прошло успешно", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else e.Cancel = true;
                    break;
                case 4:     //Number of

                    break;
                case 5:
                    table_name = "ScientificWork";

                    if (MessageBox.Show("Вы уверены, что хотите удалить данную запись. В случае наличия связанных элементов <Научный труд>, они будут также удалены!", "Внимание!!!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                    {
                        db.Delete(table_name, $"Id={displayBlock["Id", e.Row.Index].Value}");
                        db.Delete("MultiAuthorship", $"ScientificWork_Id={displayBlock["Id", e.Row.Index].Value}");
                        MessageBox.Show("Удаление прошло успешно", "Успех!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else e.Cancel = true;
                    break;
            }
        }

        private void экспортироватьВЭксельToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            try
            {
                Excel.Application excelapp = new Excel.Application();

                Excel.Workbook workbook = excelapp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.ActiveSheet;
                for (int i = 1; i < displayBlock.ColumnCount + 1; i++)
                {
                    worksheet.Rows[1].Columns[i] = displayBlock.Columns[i - 1].HeaderCell.Value;
                }
                for (int i = 2; i < displayBlock.RowCount + 2; i++)
                {
                    for (int j = 1; j < displayBlock.ColumnCount + 1; j++)
                    {
                        worksheet.Rows[i].Columns[j] = displayBlock.Rows[i - 2].Cells[j - 1].Value;
                    }
                }

                excelapp.AlertBeforeOverwriting = false;

                excelapp.Visible = true;
            }
            catch (Exception ex) { MessageBox.Show(ex.Message, "Ошибка!", MessageBoxButtons.OK, MessageBoxIcon.Stop); }
        }
        

       

        private void сформироватьВедомостьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var pd = new PrintDocument();
            // Альбомная ориентация
            pd.DefaultPageSettings.Landscape = true;
            pd.PrintPage += (s, q) =>
            {
                var bmp = new Bitmap(displayBlock.Width, displayBlock.Height);
                displayBlock.DrawToBitmap(bmp, displayBlock.ClientRectangle);
                q.Graphics.DrawImage(bmp, new Point(100, 100));
            };

            printPreviewDialog1.Document = pd;
            printPreviewDialog1.ShowDialog();
        }

        private void своднаяВедомостьПоФакультетуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bdChoose.Text = "Факультет";
            textBox7.Text = "";

            SearchBy f = new SearchBy(db, 6);
            if (f.ShowDialog() == DialogResult.OK)
            {
                string query = $"SELECT '{f.Year}' AS Период, d.Id AS Department_Id, d.Name AS DepartmentName, " +
                "sum(n.Submitted_Applications) AS Submitted_Applications, " +
                "sum(n.Submitted_Articles) AS Submitted_Articles, " +
                "sum(n.Published_Articles) AS Published_Articles, " +
                "sum(n.Abstracts) AS Abstracts " +
                "FROM NumberOf n " +
                "inner join Employee e on e.Table_Number = n.Table_Number " +
                "inner join Department d on d.Id = e.Department_Id " +
                "inner join Faculty f on f.Id = d.Faculty_Id " +
                $"WHERE f.Id = {f.Faculty} AND n.Year = {f.Year} " +
                "Group by d.Id, d.Name";

                string query1 = $"SELECT '{f.Year - 6}-{f.Year - 1}' AS Период, d.Id AS Department_Id, d.Name AS DepartmentName, " +
                   "sum(n.Submitted_Applications) AS Submitted_Applications, " +
                   "sum(n.Submitted_Articles) AS Submitted_Articles, " +
                   "sum(n.Published_Articles) AS Published_Articles, " +
                   "sum(n.Abstracts) AS Abstracts " +
                   "FROM NumberOf n " +
                   "inner join Employee e on e.Table_Number = n.Table_Number " +
                   "inner join Department d on d.Id = e.Department_Id " +
                   "inner join Faculty f on f.Id = d.Faculty_Id " +
                   $"WHERE f.Id = {f.Faculty} AND n.Year between {f.Year - 6} AND {f.Year - 1} " +
                   "Group by d.Id, d.Name";

                string query2 = $"SELECT 'До {f.Year - 6} года' AS Период, d.Id AS Department_Id, d.Name AS DepartmentName, " +
                   "sum(n.Submitted_Applications) AS Submitted_Applications, " +
                   "sum(n.Submitted_Articles) AS Submitted_Articles, " +
                   "sum(n.Published_Articles) AS Published_Articles, " +
                   "sum(n.Abstracts) AS Abstracts " +
                   "FROM NumberOf n " +
                   "inner join Employee e on e.Table_Number = n.Table_Number " +
                   "inner join Department d on d.Id = e.Department_Id " +
                   "inner join Faculty f on f.Id = d.Faculty_Id " +
                   $"WHERE f.Id = {f.Faculty} AND n.Year < {f.Year - 6} " +
                   "Group by d.Id, d.Name";

                displayBlock.AllowUserToDeleteRows = false;
                DataSet t = db.FullTable(query);
                t.Merge(db.FullTable(query1), true);
                t.Merge(db.FullTable(query2), true);
                displayBlock.DataSource = t.Tables[0];

                displayBlock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                displayBlock.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                displayBlock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

                TranslateTheTableHeader();

                label6.Text = $"Название факультета: {f.name}";
                label6.Visible = true;
            }
        }

        private void textBox7_KeyDown(object sender, KeyEventArgs e)        //Поиск по таблице
        {

            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;

                for(int i = 0; i < displayBlock.Rows.Count; i++) displayBlock.Rows[i].Visible = true;
                //displayBlock.Visible = true;
                
                string search = textBox7.Text.Trim().ToLower();

                if (string.IsNullOrEmpty(search)) return;

                if (search[0] == '!') 
                {
                    search = search.Remove(0,1);
                    for (int i = 0; i < displayBlock.Rows.Count; i++)
                    {
                        try
                        {
                            if (displayBlock["Id", i].Value.ToString().Trim().ToLower() == search ||
                                displayBlock["Table_Number", i].Value.ToString().Trim().ToLower() == search)
                            {
                                displayBlock.Rows[i].Visible = true;
                            }
                            else
                            {
                                displayBlock.CurrentCell = null;
                                displayBlock.Rows[i].Visible = false;
                            }
                        }
                        catch 
                        {
                            try
                            {
                                if (displayBlock["Id", i].Value.ToString().Trim().ToLower() == search) 
                                {
                                    displayBlock.Rows[i].Visible = true;
                                }
                                else
                                {
                                    displayBlock.CurrentCell = null;
                                    displayBlock.Rows[i].Visible = false;
                                }
                            }
                            catch { }

                            try
                            {
                                if (displayBlock["Table_Number", i].Value.ToString().Trim().ToLower() == search)
                                {
                                    displayBlock.Rows[i].Visible = true;
                                }
                                else
                                {
                                    displayBlock.CurrentCell = null;
                                    displayBlock.Rows[i].Visible = false;
                                }
                            }
                            catch { }
                        }

                    }
                }
                else
                {
                    for (int i = 0; i < displayBlock.Rows.Count; i++)
                    {

                        for (int j = 0; j < displayBlock.Columns.Count; j++)
                        {
                            if (displayBlock[j, i].Value.ToString().Trim().ToLower().Contains(search))
                            {
                                displayBlock.Rows[i].Visible = true;
                                break;
                            }
                            else
                            {
                                displayBlock.CurrentCell = null;
                                displayBlock.Rows[i].Visible = false;
                            }
                        }
                    }
                }
            }
        }

        private void ведомостьПубликацийПоКафедреToolStripMenuItem_Click(object sender, EventArgs e)
        {
            bdChoose.Text = "Кафедра";
            textBox7.Text = "";

            SearchBy f = new SearchBy(db, 8);
            if (f.ShowDialog() == DialogResult.OK)
            {
                string query = "SELECT e.FIO, e.Post, sum(n.Copyright_Certificates) AS Copyright_Certificates, " +
                "sum(n.Published_Articles) AS Published_Articles, sum(n.Abstracts) AS Abstracts " +
                "FROM Department d inner join Employee e on e.Department_Id=d.Id " +
                "inner join NumberOf n on n.Table_Number = e.Table_Number " +
                $"WHERE d.Id={f.Faculty} GROUP BY e.FIO, e.Post;";

                displayBlock.AllowUserToDeleteRows = false;
                displayBlock.DataSource = db.FullTable(query).Tables[0];

                displayBlock.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                displayBlock.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                displayBlock.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

                TranslateTheTableHeader();

                label6.Text = $"Название кафедры: {f.name}";
                label6.Visible = true;
            }
        }

        private void displayBlock_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            MessageBox.Show("Убедитесь, что вы ввели только числа", "Ошибка входных данных", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public struct Data
    {
        public List<string> name;
        public List<int> id;

        public Data(List<string> name, List<int> id)
        {
            this.name = name;
            this.id = id;
        }

        public string GetNameById(int id)
        {
            for(int i=0; i<name.Count; i++)
                if (this.id[i] == id) return name[i];
            return null;
        }
    }
}