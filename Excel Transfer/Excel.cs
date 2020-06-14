using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace Excel_Transfer {
    class Excel {

        Workbook workbook;
        Worksheet sheet;
        Dictionary<string, int> headers = new Dictionary<string, int>();

        public Excel(Application excel, string file, string worksheet) {
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

        public dynamic get(string header, int row) {
            int col = headers[header.ToLower()];
            return this.get(row, col);
        }

        public dynamic get(int row, int col) {
            return sheet.Cells[row, col].Value;
        }

        public void close() {
            workbook.Close();
        }

        public int lastRow() {
            return sheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing).Row;
        }

    }
}
