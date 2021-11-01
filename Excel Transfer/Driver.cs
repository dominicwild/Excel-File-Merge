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

        //private const string PERSON_ID = "Pers.No";
        //private const string PROJECT_ID = "Project number";
        //private const string WEEK_END = "Week Ending";

        private const string EMPLOYEE_ID = "Personnel number";
        private const string EMPLOYEE_NAME = "Name of employee or applicant";
        private const string EMPLOYEE_ROLE = "Emp Role";
        private const string EMPLOYEE_HOURLY_RATE = "Hourly Rate";
        private const string EMPLOYEE_COST_RATE = "Cost Rate";
        private const string PROJECT_ID = "Project number";
        private const string WEEK_END = "Week-End-Date";
        private const string PROJECT_TIME = "Project Time";
        private const string PROJECT_NAME = "Description";

        private const string DATE_FORMAT = "dd/MM/yyyy";
        private const string EMPLOYEE_NAME_HEADER = "Employee Name";
        private const string EMPLOYEE_ROLE_HEADER = "Employee Role";
        private const string EMPLOYEE_HOURLY_RATE_HEADER = "Hourly Rate";
        private const string EMPLOYEE_COST_RATE_HEADER = "Cost Rate";
        private const string PROJECT_NAME_HEADER = "Project Name";

        Dictionary<int, Person> people = new Dictionary<int, Person>();
        Dictionary<int, Project> projects = new Dictionary<int, Project>();
        Dictionary<string, WorkUnit> workUnits = new Dictionary<string, WorkUnit>();
        Excel labour;

        public void run(Application excel) {

            string currentDir = Directory.GetCurrentDirectory();

            string labourFile = ConfigurationManager.AppSettings["LabourFile2"];
            string labourFileWorkSheet = ConfigurationManager.AppSettings["LabourWorkSheet2"];
            string labourFileLocation = $"{currentDir}\\{labourFile}";
            string resourceTrackerFile = ConfigurationManager.AppSettings["ResourceTrackerFile"];

            log($"Loading in: {labourFileLocation} with worksheet {labourFileWorkSheet}");
            labour = new Excel(excel, labourFileLocation, labourFileWorkSheet);

            extractData();

            createSpreadSheet();
        }

        public void createSpreadSheet() {
            log("Creating the new spreadsheet.");
            Excel e = new Excel();

            string[] headers = createHeaders();
            e.setHeaders(headers);
            e.fillHeaders();

            populateSpreadsheet(e);
            formatSpreadsheet(e);

            string currentDir = Directory.GetCurrentDirectory();
            string outputFileName = ConfigurationManager.AppSettings["OutputFileName"];
            string outputFilePath = $"{currentDir}\\{outputFileName}";

            log($"Saving new spreadsheet to {outputFilePath}");
            e.saveAs(outputFilePath);
            e.close();
        }

        public void formatSpreadsheet(Excel e) {
            log("Formatting spreadsheet.");

            log("Formatting headers.");
            e.formatHeaders();

            log("Adding defaults to date headers.");
            string[] dateHeaders = e.getDateHeaders();
            e.fillWithDefault(0, dateHeaders);

            log("Adding currency formats to columns.");
            e.formatCurrency(EMPLOYEE_COST_RATE_HEADER);
            e.formatCurrency(EMPLOYEE_HOURLY_RATE_HEADER);

            log("Expanding columns.");
            e.autoExpandColumns();

            log("Formatting completed.");
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
                    fullName = $"{person.firstName}, {person.lastName}";
                } else {
                    fullName = person.firstName;
                }

                e.set(EMPLOYEE_NAME_HEADER, row, fullName);
                e.set(EMPLOYEE_COST_RATE_HEADER, row, person.costRate);
                e.set(EMPLOYEE_HOURLY_RATE_HEADER, row, person.rate);
                e.set(EMPLOYEE_ROLE_HEADER, row, person.jobRole);
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
            List<string> headers = new List<string>(new string[] { PROJECT_NAME_HEADER, EMPLOYEE_NAME_HEADER, EMPLOYEE_ROLE_HEADER, EMPLOYEE_COST_RATE_HEADER, EMPLOYEE_HOURLY_RATE_HEADER });

            headers.AddRange(dateHeaders);
            log($"Created headers: {headers}");
            return headers.ToArray();
        }

        public void extractData() {
            log($"Extracting data from spreadsheet(s).");

            int lastRow = labour.lastRow();
            int totalRows = lastRow - 1;
            int percentIncrement = (int)Math.Floor((double)lastRow / 15);

            for (int row = 2; row < lastRow; row++) {
                populatePerson(row);
                populateProject(row);
                populateWorkUnit(row);
                if (row % percentIncrement == 0) {
                    int percent = (int)Math.Round((((double)row / totalRows) * 100));
                    log($"{percent}% complete data extraction.");
                }
            }

            log($"Successfully extracted {lastRow - 1} rows of data.");
        }

        public void populateWorkUnit(int row) {

            int personId = (int)labour.get(EMPLOYEE_ID, row);
            int projectId = (int)labour.get(PROJECT_ID, row);
            string workUnitId = $"{personId}{projectId}";

            if (!workUnits.ContainsKey(workUnitId)) {
                workUnits[workUnitId] = new WorkUnit {
                    person = people[personId],
                    project = projects[projectId]
                };
            }

            WorkUnit unit = workUnits[workUnitId];
            DateTime date = (DateTime)labour.get(WEEK_END, row);
            double hoursWorked = (double)labour.get(PROJECT_TIME, row);

            unit.addWork(date, hoursWorked);

        }

        public void populatePerson(int row) {
            int id = (int)labour.get(EMPLOYEE_ID, row);
            if (!people.ContainsKey(id)) {
                Person p = new Person(id);
                string fullName = labour.get(EMPLOYEE_NAME, row);
                if (fullName.Contains(",")) {
                    p.firstName = fullName.Split(',')[1].Trim();
                    p.lastName = fullName.Split(',')[0].Trim();
                } else {
                    p.firstName = fullName;
                }

                p.jobRole = labour.get(EMPLOYEE_ROLE, row).ToString();
                p.rate = labour.get<double>(EMPLOYEE_HOURLY_RATE, row);
                p.costRate = labour.get<double>(EMPLOYEE_COST_RATE, row);



                people[id] = p;
            }
        }

        public void populateProject(int row) {
            int id = (int)labour.get(PROJECT_ID, row);
            if (!projects.ContainsKey(id)) {
                Project p = new Project(id) {
                    name = labour.get(PROJECT_NAME, row)
                };

                projects[id] = p;
            }
        }

    }
}
