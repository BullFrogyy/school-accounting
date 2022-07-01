using System;
using System.Data;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace YchetScool
{
    public partial class Autorization : Form
    {
        DB db;
        DataTable table;
        public Autorization()
        {
            InitializeComponent();
        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            Verfly();
        }

        private void Verfly()
        {
            String loginUser = loginField.Text;
            String passUser = passField.Text;

            db = new DB();
            table = new DataTable();

            MySqlDataAdapter adapter = new MySqlDataAdapter();

            MySqlCommand command = new MySqlCommand("SELECT * FROM users WHERE Login = @uL AND Password = @uP", db.GetConnection());
            command.Parameters.Add("@uL", MySqlDbType.VarChar).Value = loginUser;
            command.Parameters.Add("@uP", MySqlDbType.VarChar).Value = passUser;

            adapter.SelectCommand = command;
            adapter.Fill(table);

            if (table.Rows.Count > 0)
            {
                this.Hide();
                Form1 m = new Form1();
                m.Show();
            }
            else
                MessageBox.Show("Ошибка! Не верно введени пароль или логин!");
        }
    }
}
