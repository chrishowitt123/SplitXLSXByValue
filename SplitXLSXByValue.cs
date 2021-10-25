using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OutpaitentAnalysis
{
    class Program
    {
        static void Main(string[] args)
        {


           var dt = GetDataTableFromExcel(@"M:\MSG Open Episodes\Done\Current Episode_Final Appointment_Outcome Discharged\Copy of Current Episode_Final Appointment_Outcome Discharged_TEST1.xlsx");


            List<String> colsNames = new List<String>();
            foreach (DataColumn column in dt.Columns)
            {
                colsNames.Add(column.ToString());
            }

            var distinctGroups = dt.AsEnumerable().Select(s =>  s.Field<string>("Group")).Distinct().ToList();

            foreach(var g in distinctGroups)
            {

                Console.WriteLine(g);
            }



            DataTable dt1 = dt.Clone();

            var groups = dt.AsEnumerable().GroupBy(r => new { Group = r["Group"] });



            foreach (var g in distinctGroups)
            {
                foreach (var group in groups)
                {
                    foreach (DataRow dr in group)
                    {
                        if (dr["Group"].ToString() == g)
                        {
                            {
                                dt1.ImportRow(dr);


                            }
                        }
                    }
                }
                var xlsxFile = $@"M:\MSG Open Episodes\Done\Current Episode_Final Appointment_Outcome Discharged\Group_{g}.xlsx";
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                FileInfo fileInfo = new FileInfo(xlsxFile);
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    ExcelWorksheet ws = package.Workbook.Worksheets.Add($"Group_{g}");
                    ws.Cells["A1"].LoadFromDataTable(dt1, true);
                    ws.Cells.AutoFitColumns();
                    ws.Cells.Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    ws.View.FreezePanes(2, 1);
                    package.Save();
                    dt1.Clear();
                }
            }


            Console.WriteLine("Finished!");



















        }
        public static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets.First();
                DataTable dt = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    dt.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = dt.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column - 1] = cell.Text;
                    }
                }
                return dt;
            }
        }
    }
}

