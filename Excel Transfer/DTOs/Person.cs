using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Excel_Transfer {
    class Person {

        public string firstName { get; set; }
        public string lastName { get; set; }
        public string jobFamily { get; set; }
        public string jobRole { get; set; }
        public double rate { get; set; }
        public double costRate { get; set; }
        public int id { get; }

        public Person(int id) {
            this.id = id;
        }

    }
}
