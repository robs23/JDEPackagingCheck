using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace JDEPackagingCheck.Models
{
    public class InventorySnapshot
    {
        public int InventorySnapshotId { get; set; }
        public string ProductId { get; set; }
        public string Size { get; set; }
        public string Unit { get; set; }
        public string Status { get; set; }
        public DateTime TakenOn { get; set; }
    }
}
