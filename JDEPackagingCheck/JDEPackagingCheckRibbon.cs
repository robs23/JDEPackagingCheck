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
        public int StockColumn { get; set; } = 7;
        public int FirstColumn { get; set; } = 10;
        private void JDEPackagingCheckRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnShowCoverage_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;

            Range UsedRange = sht.UsedRange;
            foreach(Range row in UsedRange.Rows)
            {
                double stock = 0;

                if(((Range)UsedRange[row.Row, StockColumn]).Value != null)
                {
                    bool isNumber = double.TryParse(((Range)UsedRange[row.Row, StockColumn]).Value.ToString(), out stock);
                    if (isNumber)
                    {
                        for (int i = FirstColumn; i < UsedRange.Columns.Count-1; i++)
                        {
                            double currReq = 0;
                            if(((Range)UsedRange[row.Row, i]).Value != null)
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

        private void btnHideCoverage_Click(object sender, RibbonControlEventArgs e)
        {
                        Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;

            Range UsedRange = sht.UsedRange;
            foreach(Range row in UsedRange.Rows)
            {
                for (int i = FirstColumn; i < UsedRange.Columns.Count - 1; i++)
                {
                    if (((Range)UsedRange[row.Row, i]).Interior.Color == ColorTranslator.ToOle(System.Drawing.Color.LightGray))
                    {
                        ((Range)UsedRange[row.Row, i]).Interior.Color = Color.Transparent;
                    }
                }
            }
        }
    }
}
