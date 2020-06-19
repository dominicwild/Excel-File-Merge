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
using static Excel_Transfer.Logger;


namespace Excel_Transfer {
    class Program {
        [DllImport("user32.dll")]
        static extern int GetWindowThreadProcessId(int hWnd, out int lpdwProcessId);

        static void Main(string[] args) {

            //killExcelProcesses();

            log("Initialising Excel application.");
            Application excel = new Application {
                DisplayAlerts = false
            };

            int id = GetExcelProcessId(excel);
            log($"Created excel process with ID: {id}");


            Driver driver = new Driver();
            driver.run(excel);

            cleanup(excel);

            log($"The program has exited. Press any key to continue.", "Green");
            Console.ReadLine();
        }

        static void killExcelProcesses() {
            log("Killing Excel processes.");
            foreach (Process p in Process.GetProcesses()) {
                if (p.ProcessName.Equals("EXCEL")) {
                    log($"Killing excel process with id {p.Id}");
                    p.Kill();
                }
            }
            log("Killed all excel processes.");
        }

        static void cleanup(Application app) {
            int excelId = GetExcelProcessId(app);
            log($"Terminating excel process with ID: {excelId}");
            Process.GetProcessById(excelId).Kill();
            log("Starting garbage collecting.");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            Marshal.ReleaseComObject(app);
            log("Finished garbage collecting.");

        }

        public static int GetExcelProcessId(Microsoft.Office.Interop.Excel.Application excelApp) {
            int id;
            GetWindowThreadProcessId(excelApp.Hwnd, out id);
            return id;
        }
    }


}
