using Excel_Transfer.DTOs;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Transfer {
    class Driver {

        private const string PERSON_ID = "Pers.No.";
        private const string PROJECT_ID = "Project number";

        Dictionary<int, Person> people = new Dictionary<int, Person>();
        Dictionary<int, Project> projects = new Dictionary<int, Project>();
        Dictionary<string, WorkUnit> workUnits = new Dictionary<string, WorkUnit>();
        Excel labour;

        public void run(Application excel) {


            string labourFileLocation = ConfigurationManager.AppSettings["LabourFile"];
            string labourFileWorkSheet = ConfigurationManager.AppSettings["LabourWorkSheet"];
            string resourceTrackerFile = ConfigurationManager.AppSettings["ResourceTrackerFile"];

            labour = new Excel(excel, labourFileLocation, labourFileWorkSheet);

            var value = labour.get("Week Ending", 2);

            extractData();

            Excel e = new Excel("Sheet 22");
            e.set(1, 1, 34);
            e.set(4, 4, "something");
            e.set(7, 3, new DateTime());

            e.saveAs("testExcel.xlsx");
            e.close();


            Console.WriteLine(value);
            Console.WriteLine($"The variable from the config file is: {labour.lastRow()}");
            Console.ReadLine();

        }

        public void extractData() {

            int lastRow = labour.lastRow();

            for (int i = 2; i < lastRow; i++) {
                populatePerson(i);
                populateProject(i);
                populateWorkUnit(i);
            }
        }

        public void populateWorkUnit(int row) {

            int personId = (int)labour.get(PERSON_ID, row);
            int projectId = (int)labour.get(PROJECT_ID, row);
            string workUnitId = $"{personId}{projectId}";

            if (!workUnits.ContainsKey(workUnitId)) {
                workUnits[workUnitId] = new WorkUnit {
                    person = people[personId],
                    project = projects[projectId]
                };
            }

            WorkUnit unit = workUnits[workUnitId];
            DateTime date = (DateTime)labour.get("Week ending", row);
            double hoursWorked = (double)labour.get("Project time", row);

            unit.addWork(date, hoursWorked);

        }

        public void populatePerson(int row) {
            int id = (int)labour.get(PERSON_ID, row);
            if (!people.ContainsKey(id)) {
                Person p = new Person(id);
                string fullName = labour.get("Name of employee or applicant", row);
                if (fullName.Contains(",")) {
                    p.firstName = fullName.Split(',')[1].Trim();
                    p.lastName = fullName.Split(',')[0].Trim();
                } else {
                    p.firstName = fullName;
                }
                people[id] = p;
            }
        }

        public void populateProject(int row) {
            int id = (int)labour.get(PROJECT_ID, row);
            if (!projects.ContainsKey(id)) {
                Project p = new Project(id) {
                    name = labour.get("Description", row)
                };

                projects[id] = p;
            }
        }

    }
}
