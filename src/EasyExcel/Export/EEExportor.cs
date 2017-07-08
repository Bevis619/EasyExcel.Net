using EasyExcel.Extensions;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Web;
using System.Linq;
using System.IO;

namespace EasyExcel.Export
{
    /// <summary>
    /// Easy Excel Export
    /// </summary>
    public sealed class EEExportor : IDisposable
    {
        /// <summary>
        /// EPPlus 
        /// </summary>
        private ExcelPackage excelPlus;

        /// <summary>
        /// Excel Sheets
        /// </summary>
        private IList<EESheet> sheets;

        /// <summary>
        /// Excel File Name
        /// </summary>
        private string fileName;

        /// <summary>
        /// Ctor
        /// </summary>
        /// <param name="sheets">sheets</param>
        /// <param name="fileName">file name</param>
        public EEExportor(IList<EESheet> sheets, string fileName = "")
        {
            this.sheets = sheets ?? throw new ArgumentNullException(" sheets is null ");
            if (string.IsNullOrEmpty(fileName)) fileName = $"{DateTime.Now.ToDefaultDateString()}";
            this.fileName = $"{fileName}.xlsx";
        }

        /// <summary>
        /// export to straeam
        /// </summary>
        public void StreamAction()
        {
            MakeExcelAction();
            var bytes = new byte[this.excelPlus.Stream.Length];
            this.excelPlus.Stream.Read(bytes, 0, bytes.Length);
            this.Write2Response(bytes);
        }

        /// <summary>
        /// save as a file
        /// </summary>
        public bool SaveAsAction(string dir)
        {
            var isExist = Directory.Exists(dir);
            if (!isExist) return false;
            var path = Path.Combine(dir, this.fileName);
            MakeExcelAction();
            var fileInfo = new FileInfo(path);
            this.excelPlus.SaveAs(fileInfo);
            return true;
        }

        private void MakeExcelAction()
        {
            this.Dispose();
            excelPlus = new ExcelPackage();
            var tables = this.GenerateDataTable();
            foreach (var table in tables)
            {
                var sheet = excelPlus.Workbook.Worksheets.Add(table.TableName);
                var headers = this.GetHeaders(table);
                for (var i = 0; i < headers.Count(); i++)
                {
                    sheet.Cells[1, i + 1].Value = headers[i];
                }

                for (var row = 0; row < table.Rows.Count; row++)
                {
                    for (var column = 0; column < table.Columns.Count; column++)
                    {
                        sheet.Cells[row + 2, column + 1].Value = table.Rows[row][column];
                    }
                }
            }
        }

        /// <summary>
        /// Generate DataTable
        /// </summary>
        /// <returns>datatable collections</returns>
        private IList<DataTable> GenerateDataTable()
        {
            var tables = new List<DataTable>();
            var index = 1;
            foreach (var item in sheets)
            {
                var sheetTable = item.Sheets.ToDataTable();
                if (null == sheetTable) continue;
                tables.Add(sheetTable);
                if (string.IsNullOrEmpty(sheetTable.TableName)) sheetTable.TableName = item.Name;
                if (string.IsNullOrEmpty(sheetTable.TableName)) sheetTable.TableName = item.Name = $"sheet{index}";
                index++;
            }

            return tables;
        }

        /// <summary>
        /// Get Headers
        /// </summary>
        /// <param name="table">datatable</param>
        /// <returns>header collections</returns>
        private IList<string> GetHeaders(DataTable table)
        {
            var headers = new List<string>();
            if (table == null || table.Columns.Count == 0) return headers;
            foreach (DataColumn item in table.Columns)
            {
                headers.Add(item.ColumnName);
            }

            return headers;
        }

        /// <summary>
        /// Write bytes to HttpResponse
        /// </summary>
        /// <param name="bytes">bytes</param>
        private void Write2Response(byte[] bytes)
        {
            HttpContext.Current.Response.Write("<meta http-equiv=Content-Type content=text/html;charset=UTF-8>");
            HttpContext.Current.Response.ContentType = "application/vnd.ms-excel";
            HttpContext.Current.Response.AddHeader("Content-Disposition", string.Format("attachment;filename={0}", HttpUtility.UrlEncode(fileName, Encoding.UTF8)));
            HttpContext.Current.Response.Charset = "UTF-8";
            HttpContext.Current.Response.ContentEncoding = Encoding.GetEncoding("utf-8");
            HttpContext.Current.Response.Clear();
            HttpContext.Current.Response.BinaryWrite(bytes);
            HttpContext.Current.Response.End();
        }

        public void Dispose() => this.excelPlus?.Dispose();
    }
}
