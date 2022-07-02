﻿using System;
using System.Data;
using Word = Microsoft.Office.Interop.Word;
using MySql.Data.MySqlClient;
using System.Windows.Forms;

namespace YchetScool
{

    public partial class DatabaseViewer : Form
    {
        private DataTable _table;
        public MySqlConnection mycon;
        public MySqlCommand nycon;
        private string _connectData = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
        public DataSet ds;
        public string value1;
        public string value2;
        public string value3;
        public string value4;
        public string value6;
        public string value7;
        public string value8;
        public string value9;
        public string value10;
        public DatabaseViewer()
        {
            InitializeComponent();
            Initialization();
        }

        private void Initialization() 
        {
            mycon = new MySqlConnection(_connectData);
            _table = new DataTable();
        }

        private void ConnectionDatabaseClick(object sender, EventArgs e)
        {
            try
            {
                mycon.Open();
                MessageBox.Show("BD Connect");

            }
            catch 
            {
                MessageBox.Show("Connection lost");
            }
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand("SELECT ID,FIO  FROM student", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox1.DataSource = patientTable;
            comboBox1.DisplayMember = "fio";
            comboBox1.ValueMember = "id";
            DataTable patientTable2 = new DataTable();
            MySqlConnection myConnection3 = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand("SELECT ID,Title  FROM service", myConnection3);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable2);
            }
            comboBox3.DataSource = patientTable2;
            comboBox3.DisplayMember = "Title";
            comboBox3.ValueMember = "ID";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                string script = "Select student.ID as Номер,student.FIO as ФИО,student.GenderType as Пол,class.Class as " +
                    "Класс,student.Address as Адрес,student.DateOFBirth as Дата_рождения,student.Email as Почта,student.Benefits as Льготы,student.Phone as" +
                    " Телефон from student join Class on Student.Class = class.ID";
                mycon = new MySqlConnection(_connectData);
                mycon.Open();
                MySqlDataAdapter ms_data = new MySqlDataAdapter(script, _connectData);
                DataTable table = new DataTable();
                ms_data.Fill(table);
                dataGridView1.DataSource = table;
                mycon.Close();
                dataGridView1.Columns[0].Visible = false;
                const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
                DataTable patientTable = new DataTable();
                MySqlConnection myConnection = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,Class  FROM Class", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable);
                }
                DataTable patientTable2 = new DataTable();
                MySqlConnection myConnection2 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,Class  FROM Class", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable2);
                }
                comboBox25.DataSource = patientTable2;
                comboBox25.DisplayMember = "id";
                comboBox25.ValueMember = "Class";

            }
            catch
            {
                MessageBox.Show("Подключение отсутствует!");
            }

        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                string script = "Select ID as Номер,FIO as ФИО,Item as Предмет,Address as Адрес,Email as Почта,Phone as Телефон from teacher";
                MSDataFill(script, _connectData, dataGridView2);
                dataGridView2.Columns[0].Visible = false;
            }
            catch
            {
                MessageBox.Show("Подключение отсутствует");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                string script = "Select class.ID as Номер,Teacher.FIO as Классный_руководитель,class.Class as Класс,class.Cabinet as Кабинет from class join Teacher on class.ClaassRoomTeacher = Teacher.ID";
                mycon = new MySqlConnection(_connectData);
                mycon.Open();
                MySqlDataAdapter ms_data = new MySqlDataAdapter(script, _connectData);
                DataTable table = new DataTable();
                ms_data.Fill(table);
                dataGridView3.DataSource = table;
                mycon.Close();
                dataGridView3.Columns[0].Visible = false;
                const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
                DataTable patientTable = new DataTable();
                MySqlConnection myConnection = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,FIO  FROM Teacher", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable);
                }
                DataTable patientTable2 = new DataTable();
                MySqlConnection myConnection2 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,FIO  FROM Teacher", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable2);
                }
                comboBox39.DataSource = patientTable2;
                comboBox39.DisplayMember = "id";
                comboBox39.ValueMember = "fio";
            }
            catch
            {
                MessageBox.Show("Подключение отсутствует");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                string pablo = dateTimePicker13.Value.ToString("yyyy-MM-dd");
                string script = ($"INSERT INTO student(FIO,GenderType, Class,Address,Dateofbirth, Email,Benefits,Phone) values ('{textBox13.Text}','{textBox14.Text}','{comboBox25.Text}','{textBox16.Text}','{pablo}','{textBox18.Text}','{textBox19.Text}',{textBox20.Text})");
                MSDataFill(script, _connectData, dataGridView1);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            int selectedIndex = dataGridView1.SelectedRows[0].Index;
            string val = Convert.ToString(selectedIndex + 1);
            try
            {
                DialogResult dialogResult = MessageBox.Show("Вы уверены что хотите удалить договор?", "Удаление", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    string script = ($"DELETE FROM student WHERE ID = {value1} ");
                    MSDataFill(script, _connectData, dataGridView1);
                }

            }
            catch { MessageBox.Show("Неверно введены данные"); }

        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                string pablo = dateTimePicker13.Value.ToString("yyyy-MM-dd");
                string script = ($"UPDATE student SET  FIO='{textBox13.Text}',GenderType='{textBox14.Text}',Class='{comboBox25.Text}',Address='{textBox16.Text}', DateOfBirth='{pablo}',Email='{textBox18.Text}',Benefits='{textBox19.Text}',Phone={textBox20.Text} WHERE ID = {value1} ");
                MSDataFill(script, _connectData, dataGridView1);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                string script = "Select ID as Номер,Title as Название,Types as Тип,PricePerDay as Цена_за_день from service";
                MSDataFill(script, _connectData, dataGridView4);
                dataGridView4.Columns[0].Visible = false;
            }
            catch
            {
                MessageBox.Show("Подключение отсутствует");
            }
        }
        private void button11_Click_1(object sender, EventArgs e)
        {

            try
            {
                string script = ($"UPDATE service SET Title='{textBox46.Text}',Types='{textBox47.Text}',PricePerDay='{textBox48.Text}'WHERE ID = {value4}  ");
                MSDataFill(script, _connectData, dataGridView4);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                string script = ($"INSERT INTO service(Title,Types,PricePerDay) values ('{textBox46.Text}','{textBox47.Text}',{textBox48.Text})");
                MSDataFill(script, _connectData, dataGridView4);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button24_Click(object sender, EventArgs e)
        {

            try
            {
                DialogResult zz = MessageBox.Show("Вы уверены что хотите удалить договор?", "Удаление", MessageBoxButtons.YesNo);
                if (zz == DialogResult.Yes)
                {
                    string script = ($"DELETE FROM teacher WHERE ID = {value2} ");
                    MSDataFill(script, _connectData, dataGridView2);
                }


            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button21_Click(object sender, EventArgs e)
        {

            try
            {
                DialogResult zz = MessageBox.Show("Вы уверены что хотите удалить договор?", "Удаление", MessageBoxButtons.YesNo);
                if (zz == DialogResult.Yes)
                {
                    string script = ($"DELETE FROM class WHERE ID = {value3} ");
                    MSDataFill(script, _connectData, dataGridView3);
                }


            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button18_Click(object sender, EventArgs e)
        {

            try
            {


                DialogResult zz = MessageBox.Show("Вы уверены что хотите удалить договор?", "Удаление", MessageBoxButtons.YesNo);
                if (zz == DialogResult.Yes)
                {
                    string script = ($"DELETE FROM service WHERE ID = {value4}");
                    MSDataFill(script, _connectData, dataGridView4);
                }

            }
            catch { MessageBox.Show("Неверно введены данные"); }

        }

        private void button16_Click(object sender, EventArgs e)
        {
            string script = ($"DELETE FROM attendance WHERE ID = {textBox38.Text} ");
            MSDataFill(script, _connectData, dataGridView1);
        }

        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                string script = ($"UPDATE class SET  ClaassRoomTeacher={comboBox39.Text},Class='{textBox56.Text}',Cabinet={textBox57.Text} WHERE ID = {value3} ");
                MSDataFill(script, _connectData, dataGridView3);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            try
            {
                string script = ($"UPDATE teacher SET FIO='{textBox66.Text}',Item='{textBox67.Text}',Address='{textBox68.Text}',Email='{textBox69.Text}',Phone='{textBox70.Text}' WHERE ID = {value2}  ");
                MSDataFill(script, _connectData, dataGridView2);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button25_Click(object sender, EventArgs e)
        {
            try
            {
                string script = ($"INSERT INTO `teacher` (FIO,Item, Address, Email, Phone) values ('{textBox66.Text}','{textBox67.Text}','{textBox68.Text}','{textBox69.Text}','{textBox70.Text}')");
                MSDataFill(script, _connectData, dataGridView2);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button22_Click(object sender, EventArgs e)
        {
            try
            {
                string script = ($"INSERT INTO class ( ClaassRoomTeacher, Class, Cabinet) values ('{comboBox39.Text}','{textBox56.Text}','{textBox57.Text}')");
                MSDataFill(script, _connectData, dataGridView3);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            try
            {
                string script = "Select contract.ID_contract as Договор,student.FIO as Ученики,service.Title as Услуга,contract.Date_of_conclusion as Дата,contract.FIO_parents as Родители,contract.Validity_period as Период from contract  join Student on  contract.ID_student = student.ID  join service on  contract.ID_service = service.ID ";
                mycon = new MySqlConnection(_connectData);
                mycon.Open();
                MySqlDataAdapter ms_data = new MySqlDataAdapter(script, _connectData);
                DataTable table = new DataTable();
                ms_data.Fill(table);
                dataGridView7.DataSource = table;
                mycon.Close();
                dataGridView7.Columns[0].Visible = false;
                const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
                DataTable patientTable = new DataTable();
                MySqlConnection myConnection = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,FIO  FROM student", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable);
                }

                DataTable patientTable2 = new DataTable();
                MySqlConnection myConnection2 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,FIO  FROM student", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable2);
                }
                comboBox8.DataSource = patientTable2;
                comboBox8.DisplayMember = "id";
                comboBox8.ValueMember = "fio";
                DataTable patientTable3 = new DataTable();
                MySqlConnection myConnection3 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,Title  FROM service", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable3);
                }
                DataTable patientTable4 = new DataTable();
                MySqlConnection myConnection4 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,Title  FROM service", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable4);
                }
                comboBox12.DataSource = patientTable4;
                comboBox12.DisplayMember = "id";
                comboBox12.ValueMember = "Title";
            }
            catch
            {
                MessageBox.Show("Подключение отсутствует");
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            try
            {
                string script = "Select ID_lesson as Номер_занятия,Date as Дата,ID_groups as Номер_группы,Subject as Тема,Homework as Домашняя_работа,Cabinet as Кабинет from  trainingsession ";
                mycon = new MySqlConnection(_connectData);
                mycon.Open();
                MySqlDataAdapter ms_data = new MySqlDataAdapter(script, _connectData);
                DataTable table = new DataTable();
                ms_data.Fill(table);
                dataGridView8.DataSource = table;
                mycon.Close();
                dataGridView8.Columns[0].Visible = false;
                const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
                DataTable patientTable = new DataTable();
                MySqlConnection myConnection = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT Number_group  FROM  `groups`", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable);
                }
                DataTable patientTable2 = new DataTable();
                MySqlConnection myConnection2 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT Number_group  FROM  `groups`", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable2);
                }
            }
            catch
            {
                MessageBox.Show("Подключение отсутствует");
            }

        }

        private void button26_Click(object sender, EventArgs e)
        {
            try
            {
                string script = "Select attendance.ID as Номер,student.fio as Номер_ученика,attendance.DATE as Дата,attendance.Attendance as Посещения,attendance.Reason as Причина,trainingsession.subject as Номер_занятия from attendance join student on attendance.ID_student = student.ID join trainingsession on attendance.ID_traning = trainingsession.ID_lesson  ";
                mycon = new MySqlConnection(_connectData);
                mycon.Open();
                MySqlDataAdapter ms_data = new MySqlDataAdapter(script, _connectData);
                DataTable table = new DataTable();
                ms_data.Fill(table);
                dataGridView10.DataSource = table;
                mycon.Close();
                dataGridView10.Columns[0].Visible = false;
                const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
                DataTable patientTable = new DataTable();
                MySqlConnection myConnection = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,FIO  FROM student", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable);
                }
                DataTable patientTable2 = new DataTable();
                MySqlConnection myConnection2 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID,FIO  FROM student", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable2);
                }
                comboBox17.DataSource = patientTable2;
                comboBox17.DisplayMember = "id";
                comboBox17.ValueMember = "fio";

                DataTable patientTable3 = new DataTable();
                MySqlConnection myConnection3 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID_lesson,subject  FROM trainingsession", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable3);
                }
                DataTable patientTable4 = new DataTable();
                MySqlConnection myConnection4 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command = new MySqlCommand("SELECT ID_lesson,subject  FROM trainingsession", myConnection);
                    MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                    adapter.Fill(patientTable4);
                }
                comboBox21.DataSource = patientTable4;
                comboBox21.DisplayMember = "ID_lesson";
                comboBox21.ValueMember = "subject";
            }
            catch { MessageBox.Show("Подключение отсутствует"); }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            try
            {
                //string script = "Select Number_group as Номер_группы,Teacher as Учитель,ID_Service as Номер_услуги from `groups`";
                string script = "Select Number_group as Номер_группы,FIO as Учитель,Title as Номер_услуги from `groups` join Teacher on  groups.Teacher = Teacher.ID join service on  groups.ID_Service = service.ID ";
                mycon = new MySqlConnection(_connectData);
                mycon.Open();
                MySqlDataAdapter ms_data = new MySqlDataAdapter(script, _connectData);
                DataTable table = new DataTable();
                ms_data.Fill(table);
                dataGridView9.DataSource = table;
                mycon.Close();
                dataGridView9.Columns[0].Visible = false;
                const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
                DataTable patientTable = new DataTable();
                DataTable patientTable3 = new DataTable();
                MySqlConnection myConnection3 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command3 = new MySqlCommand("SELECT ID,Title  FROM service", myConnection3);
                    MySqlDataAdapter adapter3 = new MySqlDataAdapter(command3);
                    adapter3.Fill(patientTable3);
                }
                comboBox33.DataSource = patientTable3;
                comboBox33.DisplayMember = "id";
                comboBox33.ValueMember = "Title";
                DataTable patientTable4 = new DataTable();
                MySqlConnection myConnection4 = new MySqlConnection(connStr1);
                {
                    MySqlCommand command4 = new MySqlCommand("SELECT ID,FIO  FROM Teacher", myConnection4);
                    MySqlDataAdapter adapter4 = new MySqlDataAdapter(command4);
                    adapter4.Fill(patientTable4);
                }
                comboBox29.DataSource = patientTable4;
                comboBox29.DisplayMember = "id";
                comboBox29.ValueMember = "fio";

            }
            catch { MessageBox.Show("Подключение отсутствует"); }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            try
            {
                string script = ($"INSERT INTO `groups` (Teacher,ID_Service) values ('{comboBox29.Text}','{comboBox33.Text}')");
                MSDataFill(script, _connectData, dataGridView9);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult zz = MessageBox.Show("Вы уверены что хотите удалить договор?", "Удаление", MessageBoxButtons.YesNo);
                if (zz == DialogResult.Yes)
                {

                    string script = ($"DELETE FROM `groups` WHERE Number_group = {value9} ");
                    MSDataFill(script, _connectData, dataGridView9);
                }

            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            try
            {
                string script = ($"UPDATE `groups` SET  Teacher='{comboBox29.Text}',ID_Service='{comboBox33.Text}' WHERE Number_group = {value9} ");
                MSDataFill(script, _connectData, dataGridView9);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }
        private void button9_Click_1(object sender, EventArgs e)
        {
            try
            {
                string pablo = dateTimePicker10.Value.ToString("yyyy-MM-dd");
                string script = ($"UPDATE attendance SET  ID_student='{comboBox17.Text}',DATE='{pablo}',Attendance='{textBox21.Text}',Reason='{textBox11.Text}',ID_traning='{comboBox21.Text}'  WHERE ID = {value10}  ");
                MSDataFill(script, _connectData, dataGridView10);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button10_Click_1(object sender, EventArgs e)
        {
            try
            {
                DialogResult zz = MessageBox.Show("Вы уверены что хотите удалить договор?", "Удаление", MessageBoxButtons.YesNo);
                if (zz == DialogResult.Yes)
                {


                    string script = ($"DELETE FROM attendance WHERE ID = {value10} ");
                    MSDataFill(script, _connectData, dataGridView10);
                }
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button12_Click_1(object sender, EventArgs e)
        {
            try
            {
                string pablo = dateTimePicker10.Value.ToString("yyyy-MM-dd");
                string script = ($"INSERT INTO attendance ( ID_student, DATE, Attendance, Reason, ID_traning) VALUES('{comboBox17.Text}','{pablo}','{textBox21.Text}','{textBox11.Text}',{comboBox21.Text})");
                MSDataFill(script, _connectData, dataGridView4);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            try
            {
                string pablo = dateTimePicker8.Value.ToString("yyyy-MM-dd");
                string script = ($"UPDATE trainingsession SET  Date='{pablo}',ID_groups='{comboBox37.Text}',Subject='{textBox38.Text}',Homework='{textBox39.Text}',Cabinet='{textBox40.Text}'  WHERE ID_lesson = {value8}  ");
                MSDataFill(script, _connectData, dataGridView4);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult zz = MessageBox.Show("Вы уверены что хотите удалить договор?", "Удаление", MessageBoxButtons.YesNo);
                if (zz == DialogResult.Yes)
                {

                    string script = ($"DELETE FROM trainingsession WHERE ID_lesson = {value8} ");
                    MSDataFill(script, _connectData, dataGridView8);
                }
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button38_Click(object sender, EventArgs e)
        {
            try
            {
                string pablo = dateTimePicker8.Value.ToString("yyyy-MM-dd");
                string script = ($"INSERT INTO trainingsession ( Date, ID_groups, Subject, Homework, Cabinet) VALUES('{pablo}',{comboBox37.Text},'{textBox38.Text}','{textBox39.Text}',{textBox40.Text})");
                MSDataFill(script, _connectData, dataGridView8);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button40_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult zz = MessageBox.Show("Вы уверены что хотите удалить договор?", "Удаление", MessageBoxButtons.YesNo);
                if (zz == DialogResult.Yes)
                {

                    string script = ($"DELETE FROM contract WHERE ID_contract = {value7} ");
                    mycon = new MySqlConnection(_connectData);
                    MSDataFill(script, _connectData, dataGridView8);
                }
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button39_Click(object sender, EventArgs e)
        {
            try
            {
                string pablo = dateTimePicker6.Value.ToString("yyyy-MM-dd");
                string script = ($"UPDATE contract SET  ID_student='{comboBox8.Text}',ID_service='{comboBox12.Text}',Date_of_conclusion='{pablo}',FIO_parents='{textBox95.Text}',Validity_period='{textBox96.Text}'  WHERE ID_contract = {value7}  ");
                MSDataFill(script, _connectData, dataGridView7);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button41_Click(object sender, EventArgs e)
        {
            try
            {
                string pablo = dateTimePicker6.Value.ToString("yyyy-MM-dd");
                string script = ($"INSERT INTO contract (ID_student, ID_service, Date_of_conclusion, FIO_parents, Validity_period) VALUES({comboBox8.Text},{comboBox12.Text},'{pablo}','{textBox95.Text}','{textBox96.Text}')");
                MSDataFill(script, _connectData, dataGridView7);
            }
            catch { MessageBox.Show("Неверно введены данные"); }
        }

        private void button13_Click_1(object sender, EventArgs e)
        {
            string day1 = dateTimePicker1.Value.Day.ToString();
            string month1 = dateTimePicker1.Value.Month.ToString();

            string TemplateFileName = @"C:\\Users\\denis\\Documents\\Visual Studio 2015\\Projects\\YchetScool\\YchetScool\\ДОГОВОР.docx";
            var number = textBox97.Text;
            var day = day1;
            var month = month1;
            var FIO = textBox100.Text;
            var serpass = textBox101.Text;
            var numpass = textBox102.Text;
            var address = textBox103.Text;
            var child = comboBox1.Text;
            var item = comboBox3.Text;
            var datenow = dateTimePicker1.Text;
            var price = textBox107.Text;
            var pricetwo = textBox107.Text;
            var datetom = dateTimePicker2.Text;
            var FIOTWO = textBox100.Text;
            var addresstwo = textBox103.Text;
            var serpasstwo = textBox101.Text;
            var numberpass = textBox102.Text;
            var indnum = textBox115.Text;
            var datevid = dateTimePicker1.Text;
            var childtwo = comboBox1.Text;
            var wordApp = new Word.Application();
            wordApp.Visible = false;

            var wordDocument = wordApp.Documents.Open(TemplateFileName);
            ReplaceWordStub("{number}", number, wordDocument);
            ReplaceWordStub("{day}", day, wordDocument);
            ReplaceWordStub("{month}", month, wordDocument);
            ReplaceWordStub("{datevid}", datevid, wordDocument);
            ReplaceWordStub("{childtwo}", childtwo, wordDocument);
            ReplaceWordStub("{FIO}", FIO, wordDocument);
            ReplaceWordStub("{serpass}", serpass, wordDocument);
            ReplaceWordStub("{numpass}", numpass, wordDocument);
            ReplaceWordStub("{address}", address, wordDocument);
            ReplaceWordStub("{child}", child, wordDocument);
            ReplaceWordStub("{item}", item, wordDocument);
            ReplaceWordStub("{datenow}", datenow, wordDocument);
            ReplaceWordStub("{price}", price, wordDocument);
            ReplaceWordStub("{pricetwo}", pricetwo, wordDocument);
            ReplaceWordStub("{datetom}", datetom, wordDocument);
            ReplaceWordStub("{FIOTWO}", FIOTWO, wordDocument);
            ReplaceWordStub("{addresstwo}", addresstwo, wordDocument);
            ReplaceWordStub("{serpasstwo}", serpasstwo, wordDocument);
            ReplaceWordStub("{numberpass}", numberpass, wordDocument);
            ReplaceWordStub("{indnum}", indnum, wordDocument);

            wordDocument.SaveAs(@"C:\\Users\\denis\\Desktop\\result.docx");
            wordApp.Visible = true;
        }
        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT ID  FROM student WHERE FIO='{comboBox1.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox4.DataSource = patientTable;
            comboBox4.DisplayMember = "id";
            comboBox4.ValueMember = "id";
        }
        private void textBox118_TextChanged_1(object sender, EventArgs e)
        {
            string script = ($"SELECT ID_contract as Договор,student.FIO as Ученики,service.Title as Услуга,Date_of_conclusion as Дата,FIO_parents as Родители,Validity_period as Период FROM contract  join Student on  contract.ID_contract = student.ID join service on  contract.ID_contract = service.ID WHERE (((student.FIO)Like \"%" + textBox118.Text + "%\"));");
            MSDataFill(script, _connectData, dataGridView7);
        }

        private void textBox119_TextChanged(object sender, EventArgs e)
        {
            string script = ("Select trainingsession.ID_lesson as Номер_занятия,trainingsession.Date as Дата,service.Title as Номер_группы,trainingsession.Subject as Тема,trainingsession.Homework as Домашняя_работа,trainingsession.Cabinet as Кабинет from  trainingsession join service on ID_lesson = service.ID  WHERE (((service.title)Like \"%" + textBox119.Text + "%\"));");
            MSDataFill(script, _connectData, dataGridView8);
        }

        private void textBox120_TextChanged(object sender, EventArgs e)
        {
            string script = ("Select Number_group as Номер_группы,FIO as Учитель,Title as Номер_услуги from `groups` join Teacher on  `groups`.Number_group = Teacher.ID join service on  `groups`.Number_group = service.ID WHERE (((Teacher.FIO)Like \"%" + textBox120.Text + "%\"));");
            MSDataFill(script, _connectData, dataGridView9);
        }

        private void textBox121_TextChanged(object sender, EventArgs e)
        {
            string script = ("Select attendance.ID as Номер,student.fio as Номер_ученика,attendance.DATE as Дата,attendance.Attendance as Посещения,attendance.Reason as Причина,trainingsession.subject as Номер_занятия from attendance join student on attendance.ID = student.ID join trainingsession on attendance.ID = trainingsession.ID_lesson  WHERE (((student.FIO)Like \"%" + textBox121.Text + "%\"));");
            MSDataFill(script, _connectData, dataGridView10);
        }

        private void textBox123_TextChanged(object sender, EventArgs e)
        {
            string script = ("SELECT ID as Номер,Title as Название,Types as Тип,PricePerDay as Цена_за_день FROM service WHERE (((Title)Like \"%" + textBox1.Text + "%\"));");
            MSDataFill(script, _connectData, dataGridView4);
        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            string pablo = dateTimePicker3.Value.ToString("yyyy-MM-dd");
            string script = ("Select ID_contract as Договор,FIO as Ученики,Title as Услуга,Date_of_conclusion as Дата,FIO_parents as Родители,Validity_period as Период from contract  join Student on  contract.ID_contract = student.ID join service on  contract.ID_contract = service.ID WHERE (((Date_of_conclusion)Like \"%" + pablo + "%\"));");
            MSDataFill(script, _connectData, dataGridView7);
        }

        private void dateTimePicker4_ValueChanged(object sender, EventArgs e)
        {
            string pablo = dateTimePicker4.Value.ToString("yyyy-MM-dd");
            string script = ("Select trainingsession.ID_lesson as Номер_занятия,trainingsession.Date as Дата,service.Title as Номер_группы,trainingsession.Subject as Тема,trainingsession.Homework as Домашняя_работа,trainingsession.Cabinet as Кабинет from  trainingsession join service on ID_lesson = service.ID  WHERE (((Date)Like \"%" + pablo + "%\"));");
            MSDataFill(script, _connectData, dataGridView8);
        }

        private void dateTimePicker5_ValueChanged(object sender, EventArgs e)
        {
            string pablo = dateTimePicker5.Value.ToString("yyyy-MM-dd");
            string script = ("Select attendance.ID as Номер,student.fio as Номер_ученика,attendance.DATE as Дата,attendance.Attendance as Посещения,attendance.Reason as Причина,trainingsession.subject as Номер_занятия from attendance join student on attendance.ID = student.ID join trainingsession on attendance.ID = trainingsession.ID_lesson WHERE (((attendance.DATE)Like \"%" + pablo + "%\"));");
            MSDataFill(script, _connectData, dataGridView10);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT ID  FROM service WHERE Title='{comboBox3.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox5.DataSource = patientTable;
            comboBox5.DisplayMember = "id";
            comboBox5.ValueMember = "id";
        }

        private void comboBox8_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT FIO  FROM student WHERE ID='{comboBox8.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox9.DataSource = patientTable;
            comboBox9.DisplayMember = "fio";
            comboBox9.ValueMember = "fio";
        }

        private void comboBox12_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT Title  FROM service WHERE ID='{comboBox12.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox13.DataSource = patientTable;
            comboBox13.DisplayMember = "Title";
            comboBox13.ValueMember = "Title";
        }

        private void comboBox17_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT FIO  FROM student WHERE ID='{comboBox17.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox16.DataSource = patientTable;
            comboBox16.DisplayMember = "fio";
            comboBox16.ValueMember = "fio";
        }

        private void comboBox21_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT subject  FROM trainingsession WHERE ID_lesson='{comboBox21.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox20.DataSource = patientTable;
            comboBox20.DisplayMember = "subject";
            comboBox20.ValueMember = "subject";
        }

        private void comboBox25_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT Class  FROM Class WHERE ID='{comboBox25.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox24.DataSource = patientTable;
            comboBox24.DisplayMember = "Class";
            comboBox24.ValueMember = "Class";
        }

        private void comboBox29_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT FIO FROM Teacher WHERE ID='{comboBox29.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox28.DataSource = patientTable;
            comboBox28.DisplayMember = "fio";
            comboBox28.ValueMember = "fio";
        }
        private void comboBox33_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT Title  FROM service WHERE ID='{comboBox33.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox32.DataSource = patientTable;
            comboBox32.DisplayMember = "Title";
            comboBox32.ValueMember = "Title";
        }

        private void comboBox39_SelectedIndexChanged(object sender, EventArgs e)
        {
            const string connStr1 = "Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8";
            DataTable patientTable = new DataTable();
            MySqlConnection myConnection = new MySqlConnection(connStr1);
            {
                MySqlCommand command = new MySqlCommand($"SELECT FIO  FROM Teacher WHERE ID='{comboBox39.Text}'", myConnection);
                MySqlDataAdapter adapter = new MySqlDataAdapter(command);
                adapter.Fill(patientTable);
            }
            comboBox38.DataSource = patientTable;
            comboBox38.DisplayMember = "fio";
            comboBox38.ValueMember = "fio";
        }

        private void button14_Click_1(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook ExcelWorkBook;
            Microsoft.Office.Interop.Excel.Worksheet ExcelWorkSheet;
            ExcelWorkBook = ExcelApp.Workbooks.Add(System.Reflection.Missing.Value);
            ExcelWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ExcelWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < dataGridView10.Rows.Count; i++)
            {
                for (int j = 0; j < dataGridView10.ColumnCount; j++)
                {
                    ExcelApp.Cells[i + 1, j + 1] = dataGridView10.Rows[i].Cells[j].Value;
                }
            }
            ExcelApp.Visible = true;
            ExcelApp.UserControl = true;
        }

        private void UsersManualToolStripMenuItemClick(object sender, EventArgs e)
        {
            Help.ShowHelp(this, "RPend.chm");
        }

        private void AboutTheProgremmToolStripMenuItemClick(object sender, EventArgs e)
        {
            Form2 form2 = new Form2();
            form2.Show();
        }


        private void DataGridView4_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            value4 = dataGridView4.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox46.Text = dataGridView4.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox47.Text = dataGridView4.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox48.Text = dataGridView4.Rows[e.RowIndex].Cells[3].Value.ToString();
        }

        private void dataGridView3_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            value3 = dataGridView3.Rows[e.RowIndex].Cells[0].Value.ToString();
            comboBox38.Text = dataGridView3.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox56.Text = dataGridView3.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox57.Text = dataGridView3.Rows[e.RowIndex].Cells[3].Value.ToString();
        }

        private void dataGridView2_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            value2 = dataGridView2.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox66.Text = dataGridView2.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox67.Text = dataGridView2.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox68.Text = dataGridView2.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox69.Text = dataGridView2.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox70.Text = dataGridView2.Rows[e.RowIndex].Cells[5].Value.ToString();
        }

        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            value1 = dataGridView1.Rows[e.RowIndex].Cells[0].Value.ToString();
            textBox13.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString();
            textBox14.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
            comboBox24.Text = dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox16.Text = dataGridView1.Rows[e.RowIndex].Cells[4].Value.ToString();
            dateTimePicker13.Text = dataGridView1.Rows[e.RowIndex].Cells[5].Value.ToString();
            textBox18.Text = dataGridView1.Rows[e.RowIndex].Cells[6].Value.ToString();
            textBox19.Text = dataGridView1.Rows[e.RowIndex].Cells[7].Value.ToString();
            textBox20.Text = dataGridView1.Rows[e.RowIndex].Cells[8].Value.ToString();
        }

        private void dataGridView10_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            value10 = dataGridView10.Rows[e.RowIndex].Cells[0].Value.ToString();
            comboBox16.Text = dataGridView10.Rows[e.RowIndex].Cells[1].Value.ToString();
            dateTimePicker10.Text = dataGridView10.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox21.Text = dataGridView10.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox11.Text = dataGridView10.Rows[e.RowIndex].Cells[4].Value.ToString();
            comboBox20.Text = dataGridView10.Rows[e.RowIndex].Cells[5].Value.ToString();

        }

        private void dataGridView9_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            value9 = dataGridView9.Rows[e.RowIndex].Cells[0].Value.ToString();
            comboBox28.Text = dataGridView9.Rows[e.RowIndex].Cells[1].Value.ToString();
            comboBox32.Text = dataGridView9.Rows[e.RowIndex].Cells[2].Value.ToString();
        }

        private void dataGridView8_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            value8 = dataGridView8.Rows[e.RowIndex].Cells[0].Value.ToString();
            dateTimePicker8.Text = dataGridView8.Rows[e.RowIndex].Cells[1].Value.ToString();
            comboBox37.Text = dataGridView8.Rows[e.RowIndex].Cells[2].Value.ToString();
            textBox38.Text = dataGridView8.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox39.Text = dataGridView8.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox40.Text = dataGridView8.Rows[e.RowIndex].Cells[5].Value.ToString();
        }

        private void dataGridView7_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            value7 = dataGridView7.Rows[e.RowIndex].Cells[0].Value.ToString();
            comboBox9.Text = dataGridView7.Rows[e.RowIndex].Cells[1].Value.ToString();
            comboBox13.Text = dataGridView7.Rows[e.RowIndex].Cells[2].Value.ToString();
            dateTimePicker6.Text = dataGridView7.Rows[e.RowIndex].Cells[3].Value.ToString();
            textBox95.Text = dataGridView7.Rows[e.RowIndex].Cells[4].Value.ToString();
            textBox96.Text = dataGridView7.Rows[e.RowIndex].Cells[5].Value.ToString();
        }

        private void textBox1_TextChanged_1(object sender, EventArgs e)
        {
            string script = ("SELECT ID as Номер,Title as Название,Types as Тип,PricePerDay as Цена_за_день FROM service WHERE (((Title)Like \"%" + textBox1.Text + "%\"));");
            MSDataFill(script, _connectData, dataGridView4);
        }

        private void textBox2_TextChanged_1(object sender, EventArgs e)
        {
            string script = ("SELECT ID as Номер,FIO as ФИО,GenderType as Пол,Class as Класс,Address as Адрес,DateOFBirth as Дата_рождения,Email as Почта,Benefits as Льготы,Phone as Телефон FROM student WHERE (((FIO)Like \"%" + textBox2.Text + "%\"));");
            MSDataFill(script, _connectData, dataGridView1);
        }

        private void MSDataFill(string script, string connect, DataGridView dataGridView) 
        {
            mycon = new MySqlConnection(connect);
            mycon.Open();
            MySqlDataAdapter ms_data = new MySqlDataAdapter(script, connect);
            ms_data.Fill(_table);
            dataGridView.DataSource = _table;
            mycon.Close();
        }
    }
}
