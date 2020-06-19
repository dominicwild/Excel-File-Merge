using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static Excel_Transfer.Logger;

namespace Excel_Transfer {
    class Utility {

        public static double convertDouble(dynamic value) {
            double valDouble = 0;
            try {

                switch (value) {

                    case string s:
                        valDouble = Double.Parse(value);
                        break;

                    case decimal d:
                        valDouble = (double)value;
                        break;
                }

            } catch {
                log($"Could not turn {value.GetType()} with value {value} into double.", "Yellow");
            }

            return valDouble;
        }

    }
}
