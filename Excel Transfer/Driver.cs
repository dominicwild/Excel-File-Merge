using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Transfer {
    class Driver {

        Dictionary<int, Person> people = new Dictionary<int, Person>();
        Dictionary<int, Project> projects = new Dictionary<int, Project>();
        Excel labour;

        public void run(Application excel) {


            string labourFileLocation = ConfigurationManager.AppSettings["LabourFile"];
            string labourFileWorkSheet = ConfigurationManager.AppSettings["LabourWorkSheet"];
            string resourceTrackerFile = ConfigurationManager.AppSettings["ResourceTrackerFile"];

            labour = new Excel(excel, labourFileLocation, labourFileWorkSheet);

            var value = labour.get("Week Ending", 2);

            extractData();


            Console.WriteLine(value);
            Console.WriteLine($"The variable from the config file is: {labour.lastRow()}");
            Console.ReadLine();

        }

        public void extractData() {

            int lastRow = labour.lastRow();

            for (int i = 2; i < lastRow; i++) {
                populatePerson(i);
                populateProject(i);
            }
        }

        public void populatePerson(int row) {
            int id = (int)labour.get("Pers.No.", row);
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
            int id = (int)labour.get("Project number", row);
            if (!projects.ContainsKey(id)) {
                Project p = new Project(id) {
                    name = labour.get("Description", row)
                };

                projects[id] = p;
            }
        }

    }
}
