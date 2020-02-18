using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static QuangMay.Model;

namespace QuangMay
{
    class MTMLib_ToJson
    {
        public void parseXLStoJson()
        {
            Application xlApp = new Application();
            Workbook workbook =
                    xlApp.Workbooks.Open(@"\Data\MTM_Library.xlsx");
            var sheets = workbook.Sheets;
            List<MTM> listMTM = new List<MTM>();
            var SheetMinMax = new MTM();
            //foreach (var sh in sheets)
            for (int p = 1; p <= sheets.Count; p++)
            {

                //Với sheet min-max 
                //if (sh.Name == "Min-Max")
                var ActiveSheet = (Worksheet)sheets[p];
                if (ActiveSheet.Name == "Min-Max")
                {
                    for (int t = 1; t <= 10; t++)
                    {
                        for (int y = 1; y <= 11; y++)
                        {
                            var v = Convert.ToDouble(ActiveSheet.Cells[t, y].Value2);
                            v = Math.Round(v, 2);
                            SheetMinMax.sValues.Add(v);
                        }
                    }
                    SheetMinMax.sRangeMin = 0;
                    SheetMinMax.sRangeMax = 0;
                }
                else
                {
                    var thisMTM = new MTM();
                    thisMTM.sName = ActiveSheet.Name;
                    for (int t = 1; t <= 10; t++)
                    {
                        for (int y = 1; y <= 10; y++)
                        {
                            var v = Convert.ToDouble(ActiveSheet.Cells[t, y].Value2);
                            v = Math.Round(v, 2);
                            thisMTM.sValues.Add(v);
                        }
                    }


                    var cellA13 = (string)ActiveSheet.Cells[13, 1].Value;
                    if (thisMTM.sName == "MTM1")
                    {
                        thisMTM.sRangeMin = 0;
                        thisMTM.sRangeMax = Convert.ToDouble(cellA13.Split('=').Last().Trim());
                    }
                    else if (thisMTM.sName == "MTM10")
                    {
                        thisMTM.sRangeMin = Convert.ToDouble(cellA13.Split('>').Last().Trim());
                        thisMTM.sRangeMax = 1;
                    }
                    else
                    {
                        cellA13 = cellA13.Substring(7); //remove MTM for
                        thisMTM.sRangeMin = Convert.ToDouble(cellA13.Split('<').First());
                        thisMTM.sRangeMax = Convert.ToDouble(cellA13.Split('=').Last().Trim());
                    }
                    listMTM.Add(thisMTM);
                }


            }




        }
    }
}
