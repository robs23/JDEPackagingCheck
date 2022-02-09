using JDEPackagingCheck.Models;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace JDEPackagingCheck
{
    public partial class JDEPackagingCheckRibbon
    {
        public ProductKeeper productKeeper = new ProductKeeper();
        public InventorySnapshotKeeper inventorySnapshotKeeper = new InventorySnapshotKeeper();
        public DeliveryItemKeeper deliveryItemKeeper = new DeliveryItemKeeper();

        public Range UsedRange { get; set; }

        private void JDEPackagingCheckRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnShowCoverage_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;

            int? stockColumn = GetStockColumn();
            if(stockColumn != null)
            {
                int firstColumn = (int)stockColumn + 3;
                Range UsedRange = sht.UsedRange;
                string userRangeAddress = UsedRange.Address;

                foreach (Range row in UsedRange.Rows)
                {
                    double stock = 0;

                    if (((Range)UsedRange[row.Row, stockColumn]).Value != null)
                    {
                        bool isNumber = double.TryParse(((Range)UsedRange[row.Row, stockColumn]).Value.ToString(), out stock);
                        if (isNumber)
                        {
                            for (int i = firstColumn; i < UsedRange.Columns.Count - 1; i++)
                            {
                                double currReq = 0;
                                if (((Range)UsedRange[row.Row, i]).Value != null)
                                {
                                    string val = ((Range)UsedRange[row.Row, i]).Value.ToString();
                                    isNumber = double.TryParse(((Range)UsedRange[row.Row, i]).Value.ToString(), out currReq);
                                    if (isNumber)
                                    {
                                        if (currReq > stock)
                                        {
                                            break;
                                        }

                                        //decrease remaining stock
                                        stock -= currReq;
                                    }
                                }
                                //paint
                                ((Range)UsedRange[row.Row, i]).Interior.Color = Color.LightGray;
                            }
                        }
                    }
                }
            }
            else
            {
                MessageBox.Show("Nie udało się odnaleźć kolumny zawierającej zapas. Oznacz kolumnę zapasu wpisując \"z\" w nagłówku.", "Nie znalazłem zapasu..", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }

            
        }

        private void btnHideCoverage_Click(object sender, RibbonControlEventArgs e)
        {
                        Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            int? stockColumn = GetStockColumn();

            if (stockColumn != null)
            {
                int firstColumn = (int)stockColumn + 3;
                Range UsedRange = sht.UsedRange;
                foreach (Range row in UsedRange.Rows)
                {
                    for (int i = firstColumn; i < UsedRange.Columns.Count - 1; i++)
                    {
                        if (((Range)UsedRange[row.Row, i]).Interior.Color == ColorTranslator.ToOle(System.Drawing.Color.LightGray))
                        {
                            ((Range)UsedRange[row.Row, i]).Interior.Color = Color.Transparent;
                        }
                    }
                }
            }
        }

        private int? GetStockColumn()
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            int? ret = null;

            Range UsedRange = sht.UsedRange;
            foreach (Range row in UsedRange.Rows)
            {
                for (int i = 0; i < UsedRange.Columns.Count - 1; i++)
                {
                    if(((Range)UsedRange[row.Row, i]).Value != null)
                    {
                        if(((Range)UsedRange[row.Row, i]).Value.ToString().ToLower() == "z"){
                            //we found it!
                            ret = i;
                            break;
                        }
                    }
                }
                if(ret != null)
                {
                    break;
                }
            }

            return ret;
        }

        private void btnImportDeliveries_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            UsedRange = sht.UsedRange;

            bool found = false;
            int cComponent = 0;
            int cComponentName = 0;
            int cPo = 0;
            int cDocumentDate = 0;
            int cOrderQty = 0;
            int cOpenQty = 0;
            int cReceivedQty = 0;
            int cNetPrice = 0;
            int cVendor = 0;
            int cDeliveryDate = 0;

            try
            {
                for (int i = 1; i <= UsedRange.Columns.Count; i++)
                {
                    if (cComponent == 0 || cComponentName == 0 || cPo == 0 || cDocumentDate == 0 || cOrderQty == 0 || cOpenQty == 0 || cReceivedQty == 0 || cNetPrice == 0 || cVendor == 0 || cDeliveryDate == 0)
                    {

                        try
                        {
                            string aCell = ((Range)UsedRange.Cells[1, i]).Value;

                            if (aCell == "Material")
                            {
                                cComponent = i;
                            }
                            else if (aCell == "Short Text")
                            {
                                cComponentName = i;
                            }
                            else if (aCell == "Purchasing Document")
                            {
                                cPo = i;
                            }
                            else if (aCell == "Document Date")
                            {
                                cDocumentDate = i;
                            }
                            else if (aCell == "Order Quantity")
                            {
                                cOrderQty = i;
                            }
                            else if (aCell == "Still to be delivered (qty)")
                            {
                                cOpenQty = i;
                            }
                            else if (aCell == "Quantity Received")
                            {
                                cReceivedQty = i;
                            }
                            else if (aCell == "Net price")
                            {
                                cNetPrice = i;
                            }
                            else if (aCell == "Name of Vendor")
                            {
                                cVendor = i;
                            }
                            else if (aCell == "Delivery Date")
                            {
                                cDeliveryDate = i;
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                        if (cComponent != 0 && cComponentName != 0 && cPo != 0 && cDocumentDate != 0 && cOrderQty != 0 && cOpenQty != 0 && cReceivedQty != 0 && cNetPrice != 0 && cVendor != 0 && cDeliveryDate != 0)
                        {
                            found = true;
                            break;
                        }
                    }
                }

                if (found)
                {
                    //Create products not yet available in database
                    foreach(Range row in UsedRange.Rows)
                    {
                        if (((Range)UsedRange[row.Row, cComponent]).Value2 != null && ((Range)UsedRange[row.Row, cComponentName]).Value2 != null)
                        {
                            if(int.TryParse(((Range)UsedRange[row.Row, cComponent]).Value, out int ind))
                            {
                                if(ind > 0)
                                {
                                    Product p = new Product();
                                    p.ZfinIndex = ind;
                                    string name = ((Range)UsedRange[row.Row, cComponentName]).Value;
                                    p.ZfinName = name.Replace("\'", "");
                                    p.BasicUom = "PC";
                                    productKeeper.Items.Add(p);
                                }
                            }
                        }
                    }
                    productKeeper.CreateMissingProducts();
                    productKeeper.Reload();

                    //Create deliveryItems
                    foreach (Range row in UsedRange.Rows)
                    {
                        if (((Range)UsedRange[row.Row, cComponent]).Value2 != null && ((Range)UsedRange[row.Row, cPo]).Value2 != null && ((Range)UsedRange[row.Row, cDocumentDate]).Value2 != null && ((Range)UsedRange[row.Row, cOrderQty]).Value2 != null && ((Range)UsedRange[row.Row, cOpenQty]).Value2 != null && ((Range)UsedRange[row.Row, cReceivedQty]).Value2 != null && ((Range)UsedRange[row.Row, cNetPrice]).Value2 != null && ((Range)UsedRange[row.Row, cVendor]).Value2 != null && ((Range)UsedRange[row.Row, cDeliveryDate]).Value2 != null)
                        {
                            if (int.TryParse(((Range)UsedRange[row.Row, cComponent]).Value, out int ind))
                            {
                                if (ind > 0)
                                {
                                    int productId = productKeeper.Items.Where(x => x.ZfinIndex == ind).FirstOrDefault().ZfinId;
                                    if(productId > 0)
                                    {
                                        //Unrestricted stock
                                        DeliveryItem i = new DeliveryItem();
                                        i.ProductId = productId;
                                        i.DocumentDate = Convert.ToDateTime(((Range)UsedRange[row.Row, cDocumentDate]).Value);
                                        i.PurchaseOrder = ((Range)UsedRange[row.Row, cPo]).Value;
                                        double qty = 0;
                                        string sQty = ((Range)UsedRange[row.Row, cOrderQty]).Value.ToString();
                                        bool isParsable = double.TryParse(sQty, out qty);
                                        i.OrderQuantity = qty;
                                        qty = 0;
                                        sQty = ((Range)UsedRange[row.Row, cOpenQty]).Value.ToString();
                                        isParsable = double.TryParse(sQty, out qty);
                                        i.OpenQuantity = qty;
                                        qty = 0;
                                        sQty = ((Range)UsedRange[row.Row, cReceivedQty]).Value.ToString();
                                        isParsable = double.TryParse(sQty, out qty);
                                        i.ReceivedQuantity = qty;
                                        qty = 0;
                                        sQty = ((Range)UsedRange[row.Row, cNetPrice]).Value.ToString();
                                        isParsable = double.TryParse(sQty, out qty);
                                        i.NetPrice = qty;
                                        i.Vendor = ((Range)UsedRange[row.Row, cVendor]).Value;
                                        i.DeliveryDate = Convert.ToDateTime(((Range)UsedRange[row.Row, cDeliveryDate]).Value);
                                        i.CreatedOn = DateTime.Now;
                                        deliveryItemKeeper.Items.Add(i);
                                    }
                                }
                            }
                        }
                    }
                    deliveryItemKeeper.CreateSnapshot();
                    MessageBox.Show("Import zakończony powodzeniem!", "Powodzenie", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show("Nie udało się odnaleźć wszystkich kolumn.. Pawidłowy typ raportu to ME2M z SAP w układzie /M024_ZPKG", "Brakujące kolumny", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Uuups.. Coś poszło nie tak! Szczegóły: {ex.Message}", "Napotkano błędy", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnImportInventories_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            UsedRange = sht.UsedRange;

            bool found = false;
            int cComponent = 0;
            int cComponentName = 0;
            int cUnrestricted = 0;
            int cBlocked = 0;
            int cUom = 0;

            try
            {
                for (int i = 1; i <= UsedRange.Columns.Count; i++)
                {
                    if (cComponent == 0 || cComponentName == 0 || cUnrestricted == 0 || cBlocked == 0 || cUom == 0)
                    {

                        try
                        {
                            string aCell = ((Range)UsedRange.Cells[1, i]).Value;

                            if (aCell == "Material")
                            {
                                cComponent = i;
                            }
                            else if (aCell == "Material Description")
                            {
                                cComponentName = i;
                            }
                            else if (aCell == "Unrestricted")
                            {
                                cUnrestricted = i;
                            }
                            else if (aCell == "Blocked")
                            {
                                cBlocked = i;
                            }
                            else if (aCell == "Base Unit of Measure")
                            {
                                cUom = i;
                            }
                        }
                        catch (Exception ex)
                        {

                        }
                        if (cComponent != 0 && cComponentName != 0 && cUnrestricted != 0 && cBlocked != 0 & cUom != 0)
                        {
                            found = true;
                            break;
                        }
                    }
                }

                if (found)
                {
                    //Create products not yet available in database
                    foreach (Range row in UsedRange.Rows)
                    {
                        if (((Range)UsedRange[row.Row, cComponent]).Value2 != null && ((Range)UsedRange[row.Row, cComponentName]).Value2 != null && ((Range)UsedRange[row.Row, cUom]).Value2 != null)
                        {
                            if (int.TryParse(((Range)UsedRange[row.Row, cComponent]).Value, out int ind))
                            {
                                if (ind > 0)
                                {
                                    Product p = new Product();
                                    p.ZfinIndex = ind;
                                    string name = ((Range)UsedRange[row.Row, cComponentName]).Value;
                                    p.ZfinName = name.Replace("\'", "");
                                    p.BasicUom = ((Range)UsedRange[row.Row, cUom]).Value;
                                    productKeeper.Items.Add(p);
                                }
                            }
                        }
                    }
                    productKeeper.CreateMissingProducts();
                    productKeeper.Reload();

                    //Create inventorySnapshots
                    foreach (Range row in UsedRange.Rows)
                    {
                        if (((Range)UsedRange[row.Row, cComponent]).Value2 != null && ((Range)UsedRange[row.Row, cUnrestricted]).Value2 != null && ((Range)UsedRange[row.Row, cBlocked]).Value2 != null && ((Range)UsedRange[row.Row, cUom]).Value2 != null)
                        {
                            if (int.TryParse(((Range)UsedRange[row.Row, cComponent]).Value, out int ind))
                            {
                                if (ind > 0)
                                {
                                    int productId = productKeeper.Items.Where(x => x.ZfinIndex == ind).FirstOrDefault().ZfinId;
                                    if (productId > 0)
                                    {
                                        //Unrestricted stock
                                        InventorySnapshot i = new InventorySnapshot();
                                        i.ProductId = productId;
                                        i.Status = "U";
                                        double size = 0;
                                        string sSize = ((Range)UsedRange[row.Row, cUnrestricted]).Value.ToString();
                                        bool isParsable = double.TryParse(sSize, out size);
                                        i.Size = size;
                                        i.Unit = ((Range)UsedRange[row.Row, cUom]).Value;
                                        inventorySnapshotKeeper.Items.Add(i);

                                        //Blocked stock
                                        i = new InventorySnapshot();
                                        i.ProductId = productId;
                                        i.Status = "B";
                                        size = 0;
                                        sSize = ((Range)UsedRange[row.Row, cBlocked]).Value.ToString();
                                        isParsable = double.TryParse(sSize, out size);
                                        i.Size = size;
                                        i.Unit = ((Range)UsedRange[row.Row, cUom]).Value;
                                        inventorySnapshotKeeper.Items.Add(i);
                                    }
                                }
                            }
                        }
                    }
                    inventorySnapshotKeeper.CreateSnapshot();
                    MessageBox.Show("Import zakończony powodzeniem!", "Powodzenie", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
                else
                {
                    MessageBox.Show("Nie udało się odnaleźć wszystkich kolumn.. Pawidłowy typ raportu to MB52 z SAP w układzie /RVK", "Brakujące kolumny", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Uuups.. Coś poszło nie tak! Szczegóły: {ex.Message}", "Napotkano błędy", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
