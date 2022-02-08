using JDEPackagingCheck.Static;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace JDEPackagingCheck.Models
{
    public class Product
    {
        public int ZfinId { get; set; }
        public int ZfinIndex { get; set; }
        public string ZfinName { get; set; }
        public DateTime? CreationDate { get; set; }
        public DateTime? LastUpdate { get; set; }
        public string ProdStatus { get; set; }
        public string BasicUom { get; set; }

        public bool Add()
        {
            string iSql = @"INSERT INTO tbZfin (zfinIndex, zfinName, creationDate, prodStatus, basicUom)
                            output INSERTED.ZfinId 
                            VALUES(@ZfinIndex, @Name, @CreationDate, @LastUpdate, @ProdStatus, @BasicUom)";

            using (SqlCommand command = new SqlCommand(iSql, Settings.conn))
            {
                command.Parameters.AddWithValue("@ZfinIndex", ZfinIndex);
                command.Parameters.AddWithValue("@Name", ZfinName);
                command.Parameters.AddWithNullableValue("@CreationDate", CreationDate);
                command.Parameters.AddWithValue("@ProdStatus", ProdStatus);
                command.Parameters.AddWithValue("@BasicUom", BasicUom);
                int result = -1;
                try
                {
                    result = (int)command.ExecuteScalar();
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Wystąpił błąd przy dodawaniu komponentu {ZfinIndex} do bazy. Opis błędu: {ex.Message}");
                }

                if (result < 0)
                {
                    return false;
                }
                else
                {
                    ZfinId = result;
                    return true;
                }
            }
        }


    }
}
