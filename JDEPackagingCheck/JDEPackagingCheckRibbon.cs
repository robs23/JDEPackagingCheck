﻿using JDEPackagingCheck.Models;
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

        private void btnImportInventories_Click(object sender, RibbonControlEventArgs e)
        {
            Workbook wb = Globals.ThisAddIn.Application.ActiveWorkbook;
            Worksheet sht = wb.ActiveSheet;
            Range UsedRange = sht.UsedRange;

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
                    foreach(Range row in UsedRange.Rows)
                    {
                        if (((Range)UsedRange[row.Row, cComponent]).Value2 != null && ((Range)UsedRange[row.Row, cComponentName]).Value2 != null && ((Range)UsedRange[row.Row, cUom]).Value2 != null)
                        {
                            if(int.TryParse(((Range)UsedRange[row.Row, cComponent]).Value, out int ind))
                            {
                                if(ind > 0)
                                {
                                    Product p = new Product();
                                    p.ZfinIndex = ind;
                                    p.ZfinName = ((Range)UsedRange[row.Row, cComponentName]).Value;
                                    p.BasicUom = ((Range)UsedRange[row.Row, cUom]).Value;
                                    productKeeper.Items.Add(p);
                                }
                            }
                        }
                    }
                }
                else
                {
                    MessageBox.Show("Nie udało się odnaleźć wszystkich kolumn.. Pawidłowy typ raportu to MB52 z SAP w układzie /RVK", "Brakujące kolumny", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {

            }
        }
    }
}
