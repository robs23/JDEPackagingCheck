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
    }
}
