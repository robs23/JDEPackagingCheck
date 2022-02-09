using JDEPackagingCheck.Static;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDEPackagingCheck.Models
{
    public class DeliveryItemKeeper
    {
        public List<DeliveryItem> Items { get; set; }

        public DeliveryItemKeeper()
        {
            Items = new List<DeliveryItem>();
        }

        public int CreateSnapshot()
        {
            int res = -1;
            List<string> rStr = new List<string>(); //collection of deliveryItems formatted for batch upload eg (11111,'Name', 'pc'),(22222,'Name', 'kg'),... Each item contains 1000 records max (sql server requirement)
            string cStr = ""; //current item
            int counter = 0;

            foreach (DeliveryItem d in Items)
            {
                //prepare insert string
                counter++;
                if (counter % 1000 == 0)
                {
                    //we've just hit 1000 items

                    rStr.Add(cStr);
                    cStr = "";
                }
                cStr += $"({d.ProductId},'{d.DocumentDate.ToString("yyyy-MM-dd")}','{d.PurchaseOrder}',{d.OrderQuantity.ToString(CultureInfo.CreateSpecificCulture("en-GB"))},{d.OpenQuantity.ToString(CultureInfo.CreateSpecificCulture("en-GB"))},{d.ReceivedQuantity.ToString(CultureInfo.CreateSpecificCulture("en-GB"))},{d.NetPrice.ToString(CultureInfo.CreateSpecificCulture("en-GB"))},'{d.DeliveryDate.ToString("yyyy-MM-dd")}','{d.Vendor}','{d.CreatedOn.ToString("yyyy-MM-dd HH:mm:ss")}'),";

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
                    string iSql = "INSERT INTO tbDeliveryItems(ProductId, DocumentDate, PurchaseOrder, OrderQuantity, OpenQuantity, ReceivedQuantity, NetPrice, DeliveryDate, Vendor, CreatedOn) VALUES " + s;
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
