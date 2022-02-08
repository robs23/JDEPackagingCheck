using JDEPackagingCheck.Static;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDEPackagingCheck.Models
{
    public class ProductKeeper
    {
        public List<Product> Items { get; set; }

        public ProductKeeper()
        {
            Items = new List<Product>();
        }

        public void Reload()
        {

            Items.Clear();

            string sql = "SELECT zfinId, zfinIndex, zfinName, prodStatus, basicUom FROM tbZfin";

            SqlCommand sqlComand;
            sqlComand = new SqlCommand(sql, Settings.conn);
            using (SqlDataReader reader = sqlComand.ExecuteReader())
            {
                while (reader.Read())
                {
                    Product p = new Product {
                        ZfinId = reader.GetInt32(reader.GetOrdinal("zfinId")),
                        ZfinIndex = reader.GetInt32(reader.GetOrdinal("zfinIndex")),
                        ZfinName = reader["zfinName"].ToString().Trim(),
                        ProdStatus = reader["prodStatus"].ToString().Trim(),
                        BasicUom = reader["basicUom"].ToString().Trim()
                    };
                    Items.Add(p);
                }
            }
        }

        public int CreateMissingProducts()
        {
            int res = -1;
            string cSql = "CREATE TABLE #tbZfin(zfinIndex int, zfinName nvarchar(255), basicUom nvarchar(10))";
            List<string> rStr = new List<string>(); //collection of products formatted for batch upload eg (11111,'Name', 'pc'),(22222,'Name', 'kg'),... Each item contains 1000 records max (sql server requirement)
            string cStr = ""; //current item
            int counter = 0;

            using (SqlCommand command = new SqlCommand(cSql, Settings.conn))
            {
                foreach (Product p in Items)
                {
                    //prepare insert string
                    counter++;
                    if (counter % 1000 == 0)
                    {
                        //we've just hit 1000 items

                        rStr.Add(cStr);
                        cStr = "";
                    }
                    cStr += $"({p.ZfinIndex},'{p.ZfinName}','{p.BasicUom}'),";
                    
                }
                //non-full item set must be added here... otherwise it won't be added
                if (!string.IsNullOrEmpty(cStr))
                    rStr.Add(cStr);

                if (rStr.Any())
                {

                    for (int i = 0; i < rStr.Count; i++)
                    {
                        rStr[i] = rStr[i].Substring(0, rStr[i].Length - 1); //drop the last ","
                    }

                }

                command.ExecuteNonQuery();

                if (rStr.Any())
                {
                    foreach (string s in rStr)
                    {
                        //do this for each 1000 items
                        string iSql = "INSERT INTO #tbZfin(zfinIndex, zfinName, basicUom) VALUES " + s;
                        using (SqlCommand iCommand = new SqlCommand(iSql, Settings.conn))
                        {
                            iCommand.ExecuteNonQuery();
                        }
                    }

                    //once everything is uploaded to #tbZfin, differentiate it with tbZfin and add new items only
                    string sSql = $"SELECT DISTINCT zfinIndex, zfinName, GETDATE() as creationDate, 'PR' as prodStatus, basicUom FROM #tbZfin tpa WHERE NOT EXISTS (SELECT * FROM tbZfin z WHERE z.zfinIndex=tpa.zfinIndex)";
                    string iiSql = "INSERT INTO tbZfin (zfinIndex, zfinName, creationDate, prodStatus, basicUom) " + sSql;
                    using (SqlCommand iiCommand = new SqlCommand(iiSql, Settings.conn))
                    {
                        res = iiCommand.ExecuteNonQuery();
                    }
                }
            }

            return res;
        }
    }
}
