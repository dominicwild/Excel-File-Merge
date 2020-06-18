using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using static Excel_Transfer.Logger;

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
            this.excel = new Application() {
                DisplayAlerts = false
            };
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
            int col;
            try {
                col = headers[header.ToLower()];
            } catch {
                col = headers[header];
            }

            this.set(row, col, value);
        }

        public void close() {
            log($"Closing workbook {workbook.Name}");
            workbook.Close();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(sheet);
            Marshal.ReleaseComObject(workbook);
        }

        public void autoExpandColumns() {
            Range firstCell = sheet.Cells[1, 1];
            Range lastCell = sheet.Cells[this.lastRow(), this.lastColumn()];
            sheet.Range[firstCell, lastCell].Columns.AutoFit();
        }

        public void fillWithDefault(dynamic defaultValue, string[] headers) {
            for (int row = 2; row < this.lastRow() + 1; row++) {
                foreach (string header in headers) {
                    if (this.get(header, row) == null) {
                        this.set(header, row, defaultValue);
                    }
                }
            }
        }

        public string[] getDateHeaders() {
            List<string> dateHeaders = new List<string>();
            foreach (KeyValuePair<string, int> keyPair in headers) {
                string header = keyPair.Key;
                DateTime temp = new DateTime();
                if (DateTime.TryParse(header, out temp)) {
                    dateHeaders.Add(header);
                }
            }

            return dateHeaders.ToArray();
        }

        public int lastRow() {
            return sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }

        public int lastColumn() {
            return sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Column;
        }

        public void saveAs(string name) {
            workbook.SaveAs(name);
        }

        public void setHeaders(string[] headers) {
            this.headers = new Dictionary<string, int>();
            int col = 1;
            foreach (string header in headers) {
                this.headers.Add(header, col);
                col++;
            }
        }

        public void fillHeaders() {
            foreach (KeyValuePair<string, int> keyPair in headers) {
                string header = keyPair.Key;
                this.set(header, 1, header);
            }
        }

    }
}
