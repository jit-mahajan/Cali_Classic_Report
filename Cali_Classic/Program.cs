using Microsoft.Office.Interop.Excel;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using Excel = Microsoft.Office.Interop.Excel;

namespace Cali_Classic
{
    public class Program
    {

        static void Main(string[] args)
        {
            string weeklyMenumixPath = @"C:\Excel\Cali Classic\11.28 - 12.04.xlsm";
            string siteListPath = @"C:\Excel\Cali Classic\StoreList.xlsx";
            string week47path = @"C:\Excel\Cali Classic\Cali Classic Week 48.xlsm";

            Match match = Regex.Match(week47path, @"\bWeek (\d+)\b", RegexOptions.IgnoreCase);
            string numberAsString = match.Groups[1].Value;
            int week = int.Parse(numberAsString);
            week++;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            try
            {
                Excel.Application excelApp = new Excel.Application();
                excelApp.Visible = true;

                using (ExcelPackage weeklyMenumix1 = new ExcelPackage(new FileInfo(weeklyMenumixPath)))
                {
                    int count = 0;
                    for (int i = 0; i <= weeklyMenumix1.Workbook.Worksheets.Count-1; i++)
                    {
                        ExcelWorksheet worksheet = weeklyMenumix1.Workbook.Worksheets[i];

                        if (worksheet.Name == "Sheet")
                        {
                            weeklyMenumix1.Workbook.Worksheets.Delete(i);
                            Console.WriteLine("Sheet Deleted....");
                            i--;
                        }
                        else
                        {
                            Object storeValue = worksheet.Cells["G6"].Value;

                            if (storeValue != null)
                            {
                                string str = storeValue.ToString();
                                int indexOfDash = str.IndexOf('-')-1;
                                
                                if (indexOfDash >= 0)
                                {
                                    str = str.Substring(0, indexOfDash);
                                }
                               
                                worksheet.Name = str;
                                count++;
                            }
                            else
                            {
                                Console.WriteLine("Cell G6 in sheet " + worksheet.Name + " is empty.");
                            }
                        }
                    }
                    Console.WriteLine("Rename done successfully....");
                    Console.WriteLine(count);
                    weeklyMenumix1.Save();
                }

                Excel.Workbook siteList = excelApp.Workbooks.Open(siteListPath);
                Excel.Worksheet siteListsheet = (Excel.Worksheet)siteList.Sheets[1]; 
                Excel.Workbook weeklyMenumix2 = excelApp.Workbooks.Open(weeklyMenumixPath);
                Excel.Worksheet weeklyMenumixSheet = (Excel.Worksheet)weeklyMenumix2.Sheets.Add();
                weeklyMenumixSheet.Name = "Site List";

                siteListsheet.UsedRange.Copy(weeklyMenumixSheet.Range["A1"]);
                Console.WriteLine("Site List added....");


                weeklyMenumixSheet.Cells[9, 1].value = "Dist01";
                weeklyMenumixSheet.Cells[15, 1].value = "Dist03";
                weeklyMenumixSheet.Cells[22, 1].value = "Dist06";
                weeklyMenumixSheet.Cells[30, 1].value = "Dist05";
                weeklyMenumixSheet.Cells[37, 1].value = "Dist08";
                weeklyMenumixSheet.Cells[42, 1].value = "Dist02";
                weeklyMenumixSheet.Cells[49, 1].value = "Dist07";
                weeklyMenumixSheet.Cells[58, 1].value = "Dist09";
                weeklyMenumixSheet.Cells[66, 1].value = "Dist 11";
                weeklyMenumixSheet.Cells[66, 3].value = "DM";
                weeklyMenumix2.Save();
               

                Excel.Worksheet siteListSheet = (Excel.Worksheet)weeklyMenumix2.Worksheets["Site List"];
                siteListSheet.Range[$"P9:P{siteListSheet.UsedRange.Rows.Count}"].Formula = "=IF(C9=\"DM\",IF(LEFT(A9,4)=\"Dist\",CONCAT(\"D\",RIGHT(A9,2)),\" \"),P8)";
                
         
                Excel.Workbook week48 = excelApp.Workbooks.Open(week47path);
                Excel.Worksheet week48summarySheet = (Excel.Worksheet)week48.Worksheets["Summary"];
                Excel.Worksheet weeklyMenumix2Sheet = (Excel.Worksheet)weeklyMenumix2.Sheets.Add();

                weeklyMenumix2Sheet.Name = "Summary"; 
                week48summarySheet.UsedRange.Copy(weeklyMenumix2Sheet.Range["A1"]);
                weeklyMenumix2.Save();
                Console.WriteLine("Summary sheet added..");
              
                Excel.Worksheet summarySheet = (Excel.Worksheet)weeklyMenumix2.Worksheets["Summary"];
                Excel.Range usedRange = summarySheet.UsedRange;

                Excel.Range rangeToClear = usedRange.get_Range("C3:G" + usedRange.Rows.Count);
                rangeToClear.ClearContents();
                Excel.Range rangeToClear2 = usedRange.get_Range("K3:N" + Math.Min(11, usedRange.Rows.Count));
                rangeToClear2.ClearContents();
                Console.WriteLine("Sheet Cleared.....");


                summarySheet.Range[$"C3:C{summarySheet.UsedRange.Rows.Count}"].Formula = "=XLOOKUP(D3,'Site list'!A:A,'Site list'!P:P,0,0,1)";
                summarySheet.Range[$"D3:D{summarySheet.UsedRange.Rows.Count}"].Formula = "=XLOOKUP(B3,'Site list'!C:C,'Site list'!A:A,0,0,1)";
                summarySheet.Range[$"E3:E{summarySheet.UsedRange.Rows.Count}"].Formula = "=SUMIF(INDIRECT(\"'\"&$A3&\"'!D:D\"),E$1,INDIRECT(\"'\"&$A3&\"'!I:I\"))";
                summarySheet.Range[$"F3:F{summarySheet.UsedRange.Rows.Count}"].Formula = "=SUMIF(INDIRECT(\"'\"&$A3&\"'!D:D\"),F$1,INDIRECT(\"'\"&$A3&\"'!I:I\"))";
                summarySheet.Range[$"G3:G{summarySheet.UsedRange.Rows.Count}"].Formula = "=SUMIF(INDIRECT(\"'\"&$A3&\"'!D:D\"),G$1,INDIRECT(\"'\"&$A3&\"'!I:I\"))";
                summarySheet.Range[$"H3:H{summarySheet.UsedRange.Rows.Count}"].Formula = "=SUM(E3:G3)";



                int lastRow1 = Math.Min(11, summarySheet.UsedRange.Rows.Count);
                summarySheet.Range[$"K3:K{lastRow1}"].Formula = $"=SUMIF($C:$C,$J3,E:E)";

                int lastRow2 = Math.Min(11, summarySheet.UsedRange.Rows.Count);
                summarySheet.Range[$"L3:L{lastRow2}"].Formula = $"=SUMIF($C:$C,$J3,F:F)";
                
                int lastRow3 = Math.Min(11, summarySheet.UsedRange.Rows.Count);
                summarySheet.Range[$"M3:M{lastRow3}"].Formula = $"=SUMIF($C:$C,$J3,G:G)";

                int lastRow4 = Math.Min(11, summarySheet.UsedRange.Rows.Count);
                summarySheet.Range[$"N3:N{lastRow4}"].Formula = $"=SUM(K3:M3)";

                summarySheet.Rows[1].Hidden = true;
                summarySheet.Columns[1].Hidden = true;
                summarySheet.Columns[2].Hidden = true;

                Console.WriteLine("Formules added...");

                weeklyMenumix2.SaveAs($@"C:\Excel\Cali Classic\Cali Classic Week {week}.xlsm");
                Console.WriteLine("New File Created....");
                weeklyMenumix2.Close();
                week48.Close();
                siteList.Close();
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }
    }
    
}
