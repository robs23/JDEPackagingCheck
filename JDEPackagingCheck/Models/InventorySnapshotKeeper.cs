using System;
using System.Collections.Generic;
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
    }
}
