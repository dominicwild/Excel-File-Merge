using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Excel_Transfer {
    class Excel {

        Workbook workbook;
        Worksheet sheet;
        Dictionary<string, int> headers = new Dictionary<string, int>();
        Application excel;

        public Excel(Application excel, string file, string worksheet) {
            this.excel = excel;
            workbook = excel.Workbooks.Open(file);
            sheet = workbook.Worksheets[worksheet];

            int col = 1;
            var header = sheet.Cells[1, col].Value;
            if (header.GetType() == typeof(string)) {
                while (!string.IsNullOrEmpty(header)) {
                    string headerString = (string)header;
                    headerString = headerString.ToLower();
                    headers[headerString] = col;
                    col++;
                    header = sheet.Cells[1, col].Value;
                }
            }

        }

        public Excel() {
            this.excel = new Application();
            excel.SheetsInNewWorkbook = 1;
            workbook = excel.Workbooks.Add(Missing.Value);
            sheet = workbook.Worksheets[1];
        }

        public dynamic get(string header, int row) {
            try {
                int col = headers[header.ToLower()];
                return this.get(row, col);
            } catch (Exception e) {
                Console.WriteLine(e.StackTrace);
                return null;
            }
        }

        public dynamic get(int row, int col) {
            return sheet.Cells[row, col].Value;
        }

        public void set(int row, int col, dynamic value) {
            sheet.Cells[row, col] = value;
        }

        public void set(string header, int row, dynamic value) {
            int col = headers[header.ToLower()];
            this.set(row, col, value);
        }

        public void close() {
            workbook.Close();
        }

        public int lastRow() {
            return sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }

        public void saveAs(string name) {
            workbook.SaveAs(name);
        }

    }
}
