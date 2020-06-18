using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Transfer {
    class Test {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        Process GetExcelProcess(Microsoft.Office.Interop.Excel.Application excelApp) {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return Process.GetProcessById(id);
        }

        void TerminateExcelProcess(Microsoft.Office.Interop.Excel.Application excelApp) {
            var process = GetExcelProcess(excelApp);
            if (process != null) {
                process.Kill();
            }
        }


    }
}
