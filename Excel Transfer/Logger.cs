using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Transfer {
    class Logger {
        public static void log(string log, string color) {
            DateTime date = DateTime.Now;
            string dateString = date.ToString("dd/MM/yyyy H:mm:ss");
            //Console.Write($"[{dateString}]");

            ConsoleColor consoleColor;
            if (ConsoleColor.TryParse(color, out consoleColor)) {
                Console.ForegroundColor = consoleColor;
            }

            Console.WriteLine($"[{dateString}] {log}");
            Console.ResetColor();
        }

        public static void log(string log) {
            Logger.log(log, "");
        }
    }
}
