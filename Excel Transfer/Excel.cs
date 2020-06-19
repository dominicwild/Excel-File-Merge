using System;
using System.Collections.Generic;
using System.Drawing;
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
                log(e.StackTrace);
                return null;
            }
        }

        public dynamic get(int row, int col) {
            try {
                Range r = sheet.Cells[row, col];
                var a = r.NumberFormat;
                return sheet.Cells[row, col].Value;
            } catch {
                return null;
            }
        }

        public T get<T>(string header, int row) {
            return this.get<T>(row, headers[header.ToLower()]);
        }

        public T get<T>(int row, int col) {
            var value = this.get(row, col);
            object tValue = null;
            switch (typeof(T).Name.ToLower()) {
                case "double":
                    tValue = Utility.convertDouble(value);
                    break;
            }

            return (T)Convert.ChangeType(tValue, typeof(T));
        }

        public double getDouble(string header, int row) {
            return this.getDouble(row, this.headers[header]);
        }

        public double getDouble(int row, int col) {
            var value = this.get(row, col);
            double valueDouble = 0;
            if (value is string) {
                try {
                    valueDouble = Double.Parse(value);
                } catch {
                    log($"Could not turn string {value} into double at [{row},{col}].", "Yellow");
                }
            }
            return valueDouble;
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

        public void formatHeaders() {
            Range headerCells = this.getHeaderCells();
            headerCells.Interior.Color = Color.FromArgb(221, 235, 247);
            headerCells.Borders.LineStyle = XlLineStyle.xlContinuous;
            headerCells.Borders[XlBordersIndex.xlEdgeBottom].Weight = XlBorderWeight.xlThick;
            headerCells.Font.Size = 12;
        }

        public void formatCurrency(string header) {
            Range headerColumn = this.getColumn(header);
            headerColumn.NumberFormat = "$#,##0.00";
        }

        private Range getColumn(string header) {
            return this.getColumn(headers[header]);
        }

        private Range getColumn(int col) {
            int lastRow = this.lastRow();
            Range topValueCell = sheet.Cells[2, col];
            Range lastValueCell = sheet.Cells[lastRow, col];
            return sheet.Range[topValueCell, lastValueCell];
        }

        private Range getHeaderCells() {
            Range firstCell = sheet.Cells[1, 1];
            Range lastCell = sheet.Cells[1, this.lastColumn()];
            return sheet.Range[firstCell, lastCell];
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
            try {
                workbook.SaveAs(name);
            } catch {
                log($"Failed to save file {name}.", "Red");
                log($"The file may be already open in another Excel process. Close all excel processes and re-run.", "Red");
            }
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
