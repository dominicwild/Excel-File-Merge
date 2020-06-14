using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Transfer.DTOs {
    class WorkUnit {

        public Person person { get; set; }
        public Project project { get; set; }
        public Dictionary<DateTime, double> workLog { get; set; } = new Dictionary<DateTime, double>();

        public void addWork(DateTime date, double hoursWorked) {
            if (workLog.ContainsKey(date)) {
                double hours = workLog[date];
                workLog[date] = hours + hoursWorked;
            } else {
                workLog[date] = hoursWorked;
            }
        }

    }
}
