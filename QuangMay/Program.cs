using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
//using Newtonsoft.Json;
//using Newtonsoft.Json.Linq;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Web.Script.Serialization;
using static QuangMay.Model;
//using MathNet.Numerics;
//using MathNet.Numerics.Providers.LinearAlgebra;

namespace QuangMay
{
    class Program
    {
        static double b3_Tinh_HeSo_oKt(double kt)
        {
            var res = 0.16 * Math.Sin(Math.PI * kt / 0.9);
            return res;

        }

        static double b6_Tinh_Ktm(double Kt, double phiOfCity, double gocCuaGio, double xichDoVi)
        {

            double lambdaCalculus = Kt - 1.167 * Math.Pow(Kt, 3) * (1 - Kt);
            double pico = 0.979 * (1 - Kt);
            double k = 1.141 * (1 - Kt) / Kt;
            
            //todo sai sai chỗ này
            double m = 1 /
                        ((Math.Cos(phiOfCity) * Math.Cos(xichDoVi) * Math.Cos(gocCuaGio)) + 
                        (Math.Sin(phiOfCity) * Math.Sin(xichDoVi)));
            var res = lambdaCalculus + pico * Math.Exp(-k * m);
            return res;
        }

        static double b8_Tinh_X(int hour, double b7_et)
        {
            //Random rnd = new Random();
            //double et = rnd.Next(0, 6782);
            //et = et / 10000;
            //et = Math.Round(et, 4);

            double resX = dequib8(hour, b7_et);
            return resX;
        }

        static double dequib8(int hour, double et)
        {
            if (hour == 0)
            {
                return et;
            }
            hour--;
            return 0.54 * dequib8(hour, et) + et;
        }



        static double b10_Tinh_Kt(double b6_Ktm, double b9_Fnormal, double b3_oKt)
        {
            double res_kt = b6_Ktm + (b9_Fnormal * b3_oKt);
            return res_kt;
        }

        static void Main(string[] args)
        {
            Console.WriteLine("init data, please wait...");
            var goctheogio = new List<int> { -165, -150, -135, -120, -105, -90, -75, -60, -45, -30, -15, 0, 15, 30, 45, 60, 75, 90, 105, 120, 135, 150, 165, 180 };
            
            foreach (var process in Process.GetProcessesByName("EXCEL"))
            {
                process.Kill();
            }

            var ExcelFile = new FileInfo(args[0]);
            //string _pathDir = Directory.GetCurrentDirectory();
            string _pathDir = AppDomain.CurrentDomain.BaseDirectory;
            //Console.WriteLine("dir:" + _pathDir);
            var xlApp = new Application();
            var workbook = xlApp.Workbooks.Open(_pathDir + @"\Data\MTM_Library_modified.xlsx");
            var sheets = workbook.Sheets;

            var workbookAvgMonthOfCities = xlApp.Workbooks.Open(ExcelFile.FullName);
            var sheetAvgMonthOfCities = (Worksheet)workbookAvgMonthOfCities.Sheets[1];
            


            var workbookViXich = xlApp.Workbooks.Open(_pathDir + @"\Data\dovixich.xlsx");
            var Sheet_ViXich_HCM = (Worksheet)workbookViXich.Sheets[1];
            var Sheet_viXich_DANANG = (Worksheet)workbookViXich.Sheets[2];

            var errCode = "0";
            try
            {
                if (args.Count() != 1)
                {
                    errCode = "1";
                    Console.WriteLine($"Need Excel file to calculate - errCode:{errCode}");
                    return;

                }
                else
                {
                    errCode = "2";
                    var di = new DirectoryInfo(_pathDir + @"Result");
                    Directory.CreateDirectory(di.FullName);
                    errCode = "2.1";
                    foreach (FileInfo fi in di.GetFiles())
                    {
                        fi.Delete();
                    }

                    //#region init data
                    errCode = "2.2";
                    object misValue = System.Reflection.Missing.Value;
                    var KOfYear = new List<K_Year>();

                    //string path = Directory.GetCurrentDirectory();
                    //Workbook workbook = xlApp.Workbooks.Open(Server.MapPath("~/Content/MTM_Library.xlsx"));
                    //var workbook = xlApp.Workbooks.Open(Environment.CurrentDirectory + "\\Data\\MTM_Library_modified.xlsx");
                    
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
                    errCode = "2.3";




                    var phiHCM = Sheet_ViXich_HCM.Cells[1,2].Value2;
                    var arrViXichHCM = new List<double>();
                    for (int i=2;i<367;i++)
                    {
                        arrViXichHCM.Add(Sheet_ViXich_HCM.Cells[i, 4].Value2);
                    }
                    
                    var phiDANANG = Sheet_viXich_DANANG.Cells[1,2].Value2;
                    var arrViXichDANANG = new List<double>();
                    for (int i = 2; i < 367; i++)
                    {
                        arrViXichDANANG.Add(Sheet_viXich_DANANG.Cells[i, 4].Value2);
                    }
                    errCode = "3";



                    //init random number for bước 4
                    Random rnd = new Random();

                    #region parse MTM Lib to var
                    //function to create MTM_Lib.json
                    
                    List<MTM> listMTM = new List<MTM>();
                    var SheetMinMax = new MTM();
                    //foreach (var sh in sheets)
                    errCode = "4";
                    for (int p = 1; p <= sheets.Count; p++)
                    {

                        //Với sheet min-max 
                        //if (sh.Name == "Min-Max")
                        var ActiveSheet = (Worksheet)sheets[p];
                        if (ActiveSheet.Name == "Min-Max")
                        {
                            for (int t = 1; t <= 21; t++)
                            {
                                for (int y = 1; y <= 10; y++)
                                {
                                    var v = Convert.ToDouble(ActiveSheet.Cells[t, y].Value2);
                                    v = Math.Round(v, 3);
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
                                for (int y = 1; y <= 20; y++)
                                {
                                    var v = Convert.ToDouble(ActiveSheet.Cells[t, y].Value2);
                                    v = Math.Round(v, 4);
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
                        errCode = "4.1";
                        Marshal.ReleaseComObject(ActiveSheet);

                    }
                    errCode = "4.2";
                    #endregion


                    Console.WriteLine("Begin calculate...");
                    foreach (var ct in KOfYear)
                    {
                        errCode = "5";
                        List<double> TrungBinhThang = new List<double>();
                        ct.KMon.ForEach(x => TrungBinhThang.Add(x.K_AvgMon));


                        #region bước 7
                        for (int j = 0; j < ct.KMon.Count; j++)
                        {
                            #region bước 1
                            //vidu Jan 0.435
                            int indexOfMonth = j;
                            var Kthang = ct.KMon[indexOfMonth];
                            errCode = "5 b7";
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

                            errCode = "5 b7 1";
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
                            errCode = "5 b2";
                            double LastMonthAvg;
                            if (indexOfMonth == 0)
                                LastMonthAvg = ct.KMon.SingleOrDefault(x => x.MonName == "Dec").K_AvgMon;
                            else
                                //LastMonthAvg = ct.KMon[indexOfMonth - 1].K_AvgMon;
                                LastMonthAvg = ct.KMon[indexOfMonth - 1].K_DaysInMon.Last();
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
                                errCode = "5 b3";
                                var b3RowOnMinMax = 0;
                                for (int i = 0; i <= currentMinMaxColumnList.Count; i++)
                                {
                                    errCode = $"5 b3 for[i]:{i} count:{currentMinMaxColumnList.Count}";
                                    if (currentMinMaxColumnList[i] < LastDayKAvg && LastDayKAvg < currentMinMaxColumnList[i + 1])
                                    {
                                        b3RowOnMinMax = i;
                                        break;
                                    }
                                }
                                errCode = "5 b3.1";
                                double tempNo = (double)b3RowOnMinMax * 10 / 20;
                                var rowNo = Convert.ToInt16(Math.Floor(tempNo));
                                var rowOfSelectedMTM = currentMTM.sValues.GetRange(rowNo * 20, 20);
                                #endregion


                                #region bước 4
                                errCode = "5 b4";
                                var R = rnd.NextDouble();
                                R = Math.Round(R, 4);


                                #endregion

                                #region bước 5
                                errCode = "5 b5";
                                double sumOfB4 = 0;
                                int indexRowOfB4 = 0;
                                for (int i = 0; i < rowOfSelectedMTM.Count; i++)
                                {
                                    sumOfB4 += rowOfSelectedMTM[i];
                                    sumOfB4 = Math.Round(sumOfB4, 4);

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
                                errCode = "5 b6";
                                double KtOfCurrentDay = (currentMinMaxColumnList[indexRowOfB4] + currentMinMaxColumnList[indexRowOfB4 + 1]) / 2;
                                KtOfCurrentDay = Math.Round(KtOfCurrentDay, 4);
                                ct.KMon[j].K_DaysInMon[k] = KtOfCurrentDay;
                                ct.KMon[j].RndNo_DaysInMon[k] = R;
                                #endregion

                            }


                            indexOfMonth++;
                        }
                        #endregion



                        //Prepare save out result xls
                        //var idx = _path.LastIndexOf('\\');
                        //InputPath = _path.Substring(0, idx);
                        //InputPath = InputPath.Replace("UploadData", "Result");

                        errCode = "6";
                        var xlsWorkbook = xlApp.Workbooks.Add();
                        xlsWorkbook.Author = "Tôn Trương";

                        foreach (var kOfMON in ct.KMon)
                        {
                            errCode = "6 ForMon";
                            //Add a blank WorkSheet
                            //WorkSheet xlsSheet = xlsWorkbook.CreateWorkSheet(kOfMON.Key);
                            var xlsSheet = (Worksheet)xlsWorkbook.Worksheets.Add();
                            var rang = xlsSheet.UsedRange;
                            xlsSheet.Name = kOfMON.MonName;

                            xlsSheet.Cells[1][1].Value = "Day";
                            xlsSheet.Cells[1, 2].Value = "Random in step 4";
                            xlsSheet.Cells[1, 3].Value = "Result";
                            int count = 2;
                            for (int i = 1; i <= 31; i++)
                            {
                                errCode = "6 ForDay";
                                xlsSheet.Cells[count, 1].Value2 = "ngày " + i;
                                //xlsSheet.Cells[xlsSheet.Cells[count+1, 1], xlsSheet.Cells[count+24, 1]].Merge();
                                count += 25;
                                
                            }


                            count = 2;
                            for (int i = 0; i < kOfMON.K_DaysInMon.Count; i++)
                            {
                                errCode = "6 ExcelCells";
                                //xlsSheet.Cells[count, 3].Value2 = kOfMON.RndNo_DaysInMon[count];
                                //xlsSheet.Cells[count, 4].Value2 = kOfMON.K_DaysInMon[count];
                                xlsSheet.Cells[count, 2].Value2 = kOfMON.RndNo_DaysInMon[i];
                                xlsSheet.Cells[count, 3].Value2 = kOfMON.K_DaysInMon[i];

                                var b3_oKt = b3_Tinh_HeSo_oKt(kOfMON.K_DaysInMon[i]);


                                //var tongSoNgay = CountDays(31, 12);
                                for (int e = 0; e < goctheogio.Count; e++)
                                {
                                    errCode = "6 ForHour";
                                    //kOfMON index
                                    var monIndex = ct.KMon.FindIndex(g => g.MonName == kOfMON.MonName);

                                    //DaysInMon Index
                                    var countDay = CountDays(i + 1, monIndex);

                                    double phiOfCity = 0.0;
                                    double degree = 0.0;
                                    switch (ct.CityName.ToUpper())
                                    {
                                        case "TP. HCM":
                                            phiOfCity = phiHCM;
                                            degree = arrViXichHCM[countDay-1];
                                            break;
                                        case "DANANG":
                                            phiOfCity = phiDANANG;
                                            degree = arrViXichDANANG[countDay-1];
                                            break;
                                    }

                                    //==>dovixich from excel file
                                    var kt = b6_Tinh_Ktm(kOfMON.K_DaysInMon[i], phiOfCity, goctheogio[e], degree);

                                    errCode = "6 GaussFunc";
                                    //var b7_et = xlApp.WorksheetFunction.Gauss(1);
                                    var b7_et = rnd.NextDouble();
                                    //Console.WriteLine($"b7_et:{b7_et}");

                                    var x = b8_Tinh_X(e, b7_et);
                                    //var b9_fnormal = b9_fnormal();
                                    errCode = $"6 NormFunc x:{x}"; 
                                    var Fnomarl = xlApp.WorksheetFunction.Norm_S_Dist(x, true);
                                    //Fnomarl = Math.Round(Fnomarl, 5);

                                    errCode = $"6 NormFunc beforeB10";
                                    var b10 = b10_Tinh_Kt(kt, Fnomarl, b3_oKt);
                                    errCode = "6 IfHours";
                                    if (e+1 <= 6 || e+1 >= 18)
                                    {
                                        b10 = b10 > 1 ? 1 : b10;
                                    }
                                    if (b10 > 1)
                                    {
                                        b10 = b10 - Math.Truncate(b10);
                                    }
                                    xlsSheet.Cells[count + e + 1, 2].Value = $"Hour {e+1}";
                                    xlsSheet.Cells[count + e + 1, 3].Value2 = b10;
                                }


                                count+=25;
                            }

                            Marshal.ReleaseComObject(xlsSheet);
                        }
                        //Save the excel file
                        errCode = "7 SaveResultFile";
                        //Console.WriteLine($"{di}\\Result_{ct.CityName}.xlsx");
                        xlsWorkbook.SaveAs($"{di}\\Result_{ct.CityName}.xlsx");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine($"saved file to {di}\\Result_{ct.CityName}.xlsx");
                        Console.ResetColor();
                        xlsWorkbook.Close(true, misValue, misValue);

                        Marshal.ReleaseComObject(xlsWorkbook);

                    }
                    workbook.Close();
                    workbookViXich.Close();
                    workbookAvgMonthOfCities.Close();
                    //xlApp.Quit();

                    //Marshal.ReleaseComObject(sheetAvgMonthOfCities);
                    //Marshal.ReleaseComObject(sheets);
                    //Marshal.ReleaseComObject(workbook);
                    //Marshal.ReleaseComObject(workbookAvgMonthOfCities);

                    //Marshal.ReleaseComObject(Sheet_ViXich_HCM);
                    //Marshal.ReleaseComObject(Sheet_viXich_DANANG);
                    //Marshal.ReleaseComObject(workbookViXich);
                    
                    //Marshal.ReleaseComObject(xlApp);

                    Console.WriteLine("Calculate finished, check output xlsx file");
                }

            }
            catch (Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Green;
                Console.WriteLine("Error: " + e.Message);
                Console.WriteLine("ErrorCode: " + errCode);
                Console.ResetColor();
            }
            xlApp.Quit();

            Marshal.ReleaseComObject(sheetAvgMonthOfCities);
            Marshal.ReleaseComObject(sheets);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(workbookAvgMonthOfCities);

            Marshal.ReleaseComObject(Sheet_ViXich_HCM);
            Marshal.ReleaseComObject(Sheet_viXich_DANANG);
            Marshal.ReleaseComObject(workbookViXich);

            Marshal.ReleaseComObject(xlApp);
            //Console.WriteLine("any key to exit!");
            //Console.ReadLine();
        }

        static int soNgayCuaThang(int mon)
        {
            switch (mon)
            {
                case 4:
                case 6:
                case 9:
                case 11:
                    return 30;
                case 2:
                    return 28;
                default:
                    return 31;
            }
        }


        static int CountDays(int day, int mon)
        {
            var count = 0;
            for (int i= mon; i>1; i--)
            {
                count += soNgayCuaThang(i);
            }
            count += day;
            return count;
        }
    }
}
