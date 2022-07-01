using MySql.Data.MySqlClient;

namespace YchetScool
{
    class Autorizator
    {
        private MySqlConnection _connection = new MySqlConnection("Server=localhost;Database=YCHET;Uid=root;pwd=12345;charset=utf8");
        public MySqlConnection GetConnection() => _connection;

        private void OpenConnection()
        {
            if (_connection.State == System.Data.ConnectionState.Closed)
                _connection.Open();
        }
        private void CloseConnection()
        {
            if (_connection.State == System.Data.ConnectionState.Open)
                _connection.Close();
        }
    }
}
