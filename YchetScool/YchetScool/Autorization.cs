using System;
using System.Data;
using System.Windows.Forms;
using MySql.Data.MySqlClient;

namespace YchetScool
{
    public partial class Autorization : Form
    {
        private const string CMD_TEXT = "SELECT * FROM users WHERE Login = @uL AND Password = @uP";
        private Autorizator _autoriazator;
        private DataTable _table;
        private MySqlDataAdapter _adapter;
        private string _login;
        private string _password;

        public Autorization()
        {
            InitializeComponent();
            Initialization();
        }

        private void Initialization()
        {
            _autoriazator = new Autorizator();
            _table = new DataTable();
            _adapter = new MySqlDataAdapter();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Verification();
            _login = loginField.Text;
            _password = passwordField.Text;
        }

        private void Verification()
        {
            MySqlCommand command = new MySqlCommand(CMD_TEXT, _autoriazator.GetConnection());
            command.Parameters.Add("@uL", MySqlDbType.VarChar).Value = _login;
            command.Parameters.Add("@uP", MySqlDbType.VarChar).Value = _password;
            _adapter.SelectCommand = command;
            _adapter.Fill(_table);
            if (_table.Rows.Count > 0)
                OpenForm(new DatabaseViewer());
            else
                MessageBox.Show("Ошибка! Не верно введени пароль или логин!");
        }

        private void OpenForm(Form form) 
        {
            Hide();
            form.Show();
        }
    }
}
