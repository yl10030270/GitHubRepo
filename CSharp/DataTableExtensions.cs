using System;
using Microsoft.Office.Interop.Excel;

namespace Powerex.Service.FortisFtpUploader
{ 
    public static class DataTableExtensions
    {

        public static void ExportToExcel(this System.Data.DataTable dataTable, string excelFilePath = null)
        {
            try
            {
                int columnsCount;

                if (dataTable == null || (columnsCount = dataTable.Columns.Count) == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");

                // load excel, and create a new workbook
                var excel = new Application();
                excel.Workbooks.Add();

                // single worksheet
                _Worksheet worksheet = excel.ActiveSheet;

                var header = new object[columnsCount];

                // column headings               
                for (var i = 0; i < columnsCount; i++)
                    header[i] = dataTable.Columns[i].ColumnName;

                Range headerRange = worksheet.Range[(Range)worksheet.Cells[1, 1], (Range)worksheet.Cells[1, columnsCount]];
                headerRange.Value = header;
                headerRange.Font.Bold = true;

                // DataCells
                var rowsCount = dataTable.Rows.Count;
                var cells = new object[rowsCount, columnsCount];

                for (var j = 0; j < rowsCount; j++)
                    for (var i = 0; i < columnsCount; i++)
                        cells[j, i] = dataTable.Rows[j][i];

                worksheet.Range[(Range)worksheet.Cells[2, 1], (Range)worksheet.Cells[rowsCount + 1, columnsCount]].Value = cells;

                // check fielpath
                if (!string.IsNullOrEmpty(excelFilePath))
                {
                    try
                    {
                        worksheet.SaveAs(excelFilePath);
                        excel.Quit();
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("ExportToExcel: Excel file could not be saved! Check filepath.\n"
                            + ex.Message);
                    }
                }
                else
                {
                    // no filepath is given
                    excel.Visible = true;
                }
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
        }
    }
}
