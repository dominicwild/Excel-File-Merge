using Excel_Transfer.DTOs;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Excel_Transfer.Logger;

namespace Excel_Transfer {
    class Driver {

        private const string PERSON_ID = "Pers.No.";
        private const string PROJECT_ID = "Project number";
        private const string DATE_FORMAT = "dd/MM/yyyy";
        private const string EMPLOYEE_NAME_HEADER = "Employee Name";
        private const string PROJECT_NAME_HEADER = "Project Name";

        Dictionary<int, Person> people = new Dictionary<int, Person>();
        Dictionary<int, Project> projects = new Dictionary<int, Project>();
        Dictionary<string, WorkUnit> workUnits = new Dictionary<string, WorkUnit>();
        Excel labour;

        public void run(Application excel) {

            string currentDir = Directory.GetCurrentDirectory();

            string labourFileLocation = ConfigurationManager.AppSettings["LabourFile"];
            string labourFileWorkSheet = ConfigurationManager.AppSettings["LabourWorkSheet"];
            string resourceTrackerFile = ConfigurationManager.AppSettings["ResourceTrackerFile"];

            log($"Loading in: {labourFileLocation} with worksheet {labourFileWorkSheet}");
            labour = new Excel(excel, labourFileLocation, labourFileWorkSheet);

            extractData();

            createSpreadSheet();



            Console.WriteLine($"The current working directory is: {currentDir}");
            Console.ReadLine();

        }

        public void createSpreadSheet() {
            log("Creating the new spreadsheet.");
            Excel e = new Excel();

            string[] headers = createHeaders();
            e.setHeaders(headers);
            e.fillHeaders();

            populateSpreadsheet(e);
            formatSpreadsheet(e);

            var a = e.get(5, 5);

            string outputDir = ConfigurationManager.AppSettings["OutputDirectory"];

            log("Saving new spreadsheet.");
            e.saveAs($"{outputDir}testExcel2.xlsx");
            e.close();
        }

        public void formatSpreadsheet(Excel e) {
            e.autoExpandColumns();
            string[] dateHeaders = e.getDateHeaders();
            e.fillWithDefault(0, dateHeaders);
        }

        public void populateSpreadsheet(Excel e) {
            log("Populating new spreadsheet.");
            int row = 2;
            foreach (KeyValuePair<string, WorkUnit> keyPair in workUnits) {
                log($"Writing row {row}.");
                WorkUnit unit = keyPair.Value;
                Person person = unit.person;
                Project project = unit.project;
                string fullName = "";
                if (!String.IsNullOrEmpty(person.lastName)) {
                    fullName = $"{person.firstName} , {person.lastName}";
                } else {
                    fullName = person.firstName;
                }

                e.set(EMPLOYEE_NAME_HEADER, row, fullName);
                e.set(PROJECT_NAME_HEADER, row, project.name);

                foreach (KeyValuePair<DateTime, double> logKeyPair in unit.workLog) {
                    string dateHeader = logKeyPair.Key.ToString(DATE_FORMAT);
                    double timeSpent = logKeyPair.Value;
                    e.set(dateHeader, row, timeSpent);
                }

                row++;
            }
            log($"Finished writing to spreadsheet.");
        }

        public string[] createHeaders() {
            log($"Creating the spreadsheet headers.");
            HashSet<DateTime> dateSet = new HashSet<DateTime>();

            log($"Gathering date headers.");
            foreach (KeyValuePair<string, WorkUnit> keyPair in workUnits) {
                WorkUnit unit = keyPair.Value;
                foreach (KeyValuePair<DateTime, double> date in unit.workLog) {
                    dateSet.Add(date.Key);
                }
            }

            DateTime[] dates = dateSet.ToArray();
            Array.Sort(dates);

            string[] dateHeaders = dates.Select(date => date.ToString(DATE_FORMAT)).ToArray();
            List<string> headers = new List<string>(new string[] { PROJECT_NAME_HEADER, EMPLOYEE_NAME_HEADER });

            headers.AddRange(dateHeaders);
            log($"Created headers: {headers}");
            return headers.ToArray();
        }

        public void extractData() {
            log($"Extracing data from spreadsheet(s).");

            int lastRow = labour.lastRow();

            for (int i = 2; i < lastRow; i++) {
                populatePerson(i);
                populateProject(i);
                populateWorkUnit(i);
            }

            log($"Successfully extracted {lastRow - 1} rows of data.");
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
