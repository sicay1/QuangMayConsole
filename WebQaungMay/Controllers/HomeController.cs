using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web;
using System.Web.Mvc;
using static WebQaungMay.Models.Model;

namespace WebQaungMay.Controllers
{
    public class HomeController : Controller
    {
        [HttpGet]
        public ActionResult Index()
        {

            return View();
        }


        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            try
            {
                string _path = "";
                if (file.ContentLength > 0)
                {
                    string _FileName = Path.GetFileName(file.FileName);
                    _path = Path.Combine(Server.MapPath("~/UploadData"), _FileName);
                    var ExcelFile = new FileInfo(_path);
                    if(ExcelFile.Exists)
                        ExcelFile.Delete();
                    file.SaveAs(_path);
                }
                //ViewBag.Message = "File Uploaded Successfully!!";



                //#region init data
                Application xlApp = new Application();
                //var xlsWorkbook = new Workbook();

                var InputPath = "";
                object misValue = System.Reflection.Missing.Value;
                var KOfYear = new List<K_Year>();


                string path = Directory.GetCurrentDirectory();

                Workbook workbook = xlApp.Workbooks.Open(Server.MapPath("~/Content/MTM_Library.xlsx"));

                Workbook workbookAvgMonthOfCities = xlApp.Workbooks.Open(_path);

                var sheetAvgMonthOfCities = (Worksheet)workbookAvgMonthOfCities.Sheets[1];
                int CurrColumn = 0;

                while (!string.IsNullOrEmpty(sheetAvgMonthOfCities.Cells[1, CurrColumn + 2].Text))
                {
                    KOfYear.Add(new K_Year());
                    KOfYear[CurrColumn].CityName = sheetAvgMonthOfCities.Cells[1, CurrColumn + 2].Text;
                    for (int i = 0; i < 12; i++)
                    {
                        KOfYear[CurrColumn].KMon[i].K_AvgMon = sheetAvgMonthOfCities.Cells[i + 2, CurrColumn + 2].Value2;
                    }
                    CurrColumn++;
                }

                //init random number for bước 4
                Random rnd = new Random();

                #region parse MTM Lib to var
                //function to create MTM_Lib.json
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




                //FileInfo fi1 = new FileInfo(@path + @"\Data\MTM_MinMax.json");
                //StreamReader sr = new StreamReader(File.ReadAllText(fi1.FullName));
                //string jsonString = sr.ReadToEnd();
                //JavaScriptSerializer ser = new JavaScriptSerializer();
                //List<MTM> listMTM = ser.Deserialize<List<MTM>>(jsonString);

                //FileInfo fi2 = new FileInfo(@path + @"\Data\MTM_MinMax.json");
                //sr = new StreamReader(File.ReadAllText(fi2.FullName));
                //jsonString = sr.ReadToEnd();
                //MTM SheetMinMax = ser.Deserialize<MTM>(jsonString);


                #endregion


                //var aaa = JsonConvert.SerializeObject(SheetMinMax);




                foreach (var ct in KOfYear)
                {
                    List<double> TrungBinhThang = new List<double> {
                            0.435, 0.345, 0.475, 0.550, 0.510, 0.500, 0.520, 0.498, 0.485, 0.475, 0.464, 0.452,
                            };






                    #region bước 7
                    for (int j = 0; j < ct.KMon.Count; j++)
                    {
                        #region bước 1
                        //vidu Jan 0.435
                        int indexOfMonth = j;
                        var Kthang = ct.KMon[indexOfMonth];

                        //lấy range tat ca cac sheet để so sánh
                        //vidu vậy chọn Sheet “MTM 4” trong file Excel(0.40 < Kt ≤ 0.45)
                        var currentMTM = new MTM();
                        foreach (var mtm in listMTM)
                        {
                            if (mtm.sRangeMin < Kthang.K_AvgMon && Kthang.K_AvgMon <= mtm.sRangeMax)
                            {
                                currentMTM = mtm;
                                break;
                            }
                        }

                        var temp = currentMTM.sName.Substring(3);
                        var currentMinMaxColumn = Convert.ToInt16(temp);
                        var currentMinMaxColumnList = new List<double>();

                        for (int i = 0; i < SheetMinMax.sValues.Count; i = i + 10)
                        {
                            if (SheetMinMax.sValues.Count < i + currentMinMaxColumn - 1)
                                break;

                            currentMinMaxColumnList.Add(SheetMinMax.sValues[i + currentMinMaxColumn - 1]);

                        }
                        #endregion

                        #region bước 2
                        double LastMonthAvg;
                        if (indexOfMonth == 0)
                            LastMonthAvg = ct.KMon.SingleOrDefault(x => x.MonName == "Dec").K_AvgMon;
                        else
                            LastMonthAvg = ct.KMon[indexOfMonth - 1].K_AvgMon;
                        #endregion

                        //loop for days in current month
                        for (int k = 0; k < ct.KMon[j].K_DaysInMon.Capacity; k++)
                        {
                            double LastDayKAvg = 0;
                            if (k == 0)
                            {
                                LastDayKAvg = LastMonthAvg;
                            }
                            else
                            {
                                LastDayKAvg = ct.KMon[j].K_DaysInMon[k - 1];
                            }
                            #region bước 3
                            var b3RowOnMinMax = 0;
                            for (int i = 0; i <= currentMinMaxColumnList.Count; i++)
                            {
                                if (currentMinMaxColumnList[i] < LastDayKAvg && LastDayKAvg < currentMinMaxColumnList[i + 1])
                                {
                                    b3RowOnMinMax = i;
                                    break;
                                }
                            }

                            var rowNo = b3RowOnMinMax * 10;
                            var rowOfSelectedMTM = currentMTM.sValues.GetRange(rowNo, 10);
                            #endregion


                            #region bước 4

                            var R = rnd.NextDouble();
                            R = Math.Round(R, 2);


                            #endregion

                            #region bước 5
                            double sumOfB4 = 0;
                            int indexRowOfB4 = 0;
                            for (int i = 0; i < rowOfSelectedMTM.Count - 1; i++)
                            {
                                sumOfB4 += rowOfSelectedMTM[i];
                                sumOfB4 = Math.Round(sumOfB4, 3);

                                if (sumOfB4 > R)
                                {
                                    indexRowOfB4 = i;
                                    break;
                                }
                            }

                            //double CurrentKt = 0;
                            //for (int i=0; i<currentMinMaxColumnList.Count-1;i++)
                            //{
                            //    if (currentMinMaxColumnList[i+1] > sumOfB4)
                            //    {
                            //        CurrentKt = (currentMinMaxColumnList[i] + currentMinMaxColumnList[i + 1]) / 2;

                            //    }
                            //}
                            #endregion

                            #region bước 6
                            double KtOfCurrentDay = (currentMinMaxColumnList[indexRowOfB4] + currentMinMaxColumnList[indexRowOfB4 + 1]) / 2;
                            KtOfCurrentDay = Math.Round(KtOfCurrentDay, 2);
                            ct.KMon[j].K_DaysInMon[k] = KtOfCurrentDay;
                            ct.KMon[j].RndNo_DaysInMon[k] = R;
                            #endregion

                        }


                        indexOfMonth++;
                    }
                    #endregion



                    //Prepare save out result xls
                    var idx = _path.LastIndexOf('\\');
                    InputPath = _path.Substring(0, idx);

                    var xlsWorkbook = xlApp.Workbooks.Add();
                    xlsWorkbook.Author = "Tôn Trương";

                    foreach (var kOfMON in ct.KMon)
                    {
                        //Add a blank WorkSheet
                        //WorkSheet xlsSheet = xlsWorkbook.CreateWorkSheet(kOfMON.Key);
                        var xlsSheet = (Microsoft.Office.Interop.Excel.Worksheet)xlsWorkbook.Worksheets.Add();
                        var rang = xlsSheet.UsedRange;
                        xlsSheet.Name = kOfMON.MonName;

                        xlsSheet.Cells[1][1].Value = "Day";
                        xlsSheet.Cells[1, 2].Value = "Random in step 4";
                        xlsSheet.Cells[1, 3].Value = "Result";
                        int count = 2;
                        for (int i = 1; i <= 31; i++)
                        {
                            xlsSheet.Cells[count, 1].Value2 = i;
                            count++;
                        }


                        count = 2;
                        for (int i = 0; i < kOfMON.K_DaysInMon.Count; i++)
                        {
                            //xlsSheet.Cells[count, 3].Value2 = kOfMON.RndNo_DaysInMon[count];
                            //xlsSheet.Cells[count, 4].Value2 = kOfMON.K_DaysInMon[count];
                            xlsSheet.Cells[count, 2].Value2 = kOfMON.RndNo_DaysInMon[i];
                            xlsSheet.Cells[count, 3].Value2 = kOfMON.K_DaysInMon[i];
                            count++;
                        }

                        Marshal.ReleaseComObject(xlsSheet);
                    }
                    //Save the excel file
                    xlsWorkbook.SaveAs($"{InputPath}\\Result_{ct.CityName}.xlsx");
                    Console.WriteLine($"saved file to {InputPath}\\Result_{ct.CityName}.xlsx");
                    xlsWorkbook.Close(true, misValue, misValue);

                    Marshal.ReleaseComObject(xlsWorkbook);










                }
                workbook.Close();

                workbookAvgMonthOfCities.Close();
                xlApp.Quit();


                Marshal.ReleaseComObject(sheetAvgMonthOfCities);
                Marshal.ReleaseComObject(sheets);

                Marshal.ReleaseComObject(xlApp);
                Marshal.ReleaseComObject(workbook);
                Marshal.ReleaseComObject(workbookAvgMonthOfCities);

                ViewBag.Message = "okkkkkkkkk";
                return View();

            }
            catch(Exception e)
            {
                ViewBag.Message = "File upload failed!! " + e.Message;
                return View();
            }
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}