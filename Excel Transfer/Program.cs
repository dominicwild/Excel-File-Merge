using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Transfer {
    class Program {
        static void Main(string[] args) {

            killExcelProcesses();

            Application excel = new Application {
                DisplayAlerts = false
            };

            Driver driver = new Driver();
            driver.run(excel);

            //book.SaveAs("test24.xlsx");
            //book.Close(true, $"{currentDirectory}\\test.xlsx");
            //book.Close();

            cleanup(excel);
        }

        static void killExcelProcesses() {
            foreach (Process p in Process.GetProcesses()) {
                if (p.ProcessName.Equals("EXCEL")) {
                    p.Kill();
                }
            }
        }

        static void cleanup(Application app) {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(app);
        }
    }


}
