using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Transfer {
    class Project {

        public int id { get; set; }
        public string name { get; set; }
        public string wbsCode { get; set; }

        public Project(int id) {
            this.id = id;
        }
    }
}
