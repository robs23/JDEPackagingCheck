using JDEPackagingCheck.Static;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDEPackagingCheck.Models
{
    public class InventorySnapshotKeeper
    {
        public List<InventorySnapshot> Items { get; set; }

        public InventorySnapshotKeeper()
        {
            Items = new List<InventorySnapshot>();
        }

        public int CreateSnapshot()
        {
            int res = -1;
            List<string> rStr = new List<string>(); //collection of inventories formatted for batch upload eg (11111,'Name', 'pc'),(22222,'Name', 'kg'),... Each item contains 1000 records max (sql server requirement)
            string cStr = ""; //current item
            int counter = 0;

            foreach (InventorySnapshot i in Items)
            {
                //prepare insert string
                counter++;
                if (counter % 1000 == 0)
                {
                    //we've just hit 1000 items

                    rStr.Add(cStr);
                    cStr = "";
                }
                cStr += $"({i.ProductId},{i.Size},'{i.Unit}','{DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")}','{i.Status}'),";

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

            if (rStr.Any())
            {
                foreach (string s in rStr)
                {
                    //do this for each 1000 items
                    string iSql = "INSERT INTO tbInventorySnapshots(ProductId, Size, Unit, TakenOn, Status) VALUES " + s;
                    using (SqlCommand iCommand = new SqlCommand(iSql, Settings.conn))
                    {
                        res = iCommand.ExecuteNonQuery();
                    }
                }
            }
            

            return res;
        }

    }
}
