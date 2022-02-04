using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JDEPackagingCheck.Static
{
    public static class Settings
    {
        private static SqlConnection _conn { get; set; }
        public static SqlConnection conn
        {
            get
            {
                if (_conn == null)
                {
                    _conn = new SqlConnection(Static.Secrets.ConnectionString);
                }
                if (_conn.State == System.Data.ConnectionState.Closed || _conn.State == System.Data.ConnectionState.Closed)
                {
                    try
                    {
                        _conn.Open();
                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show("Nie udało się nawiązać połączenia z bazą danych.. " + ex.Message);
                    }

                }
                return _conn;
            }
        }
    }
}
