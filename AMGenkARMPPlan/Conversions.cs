using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AMGenkARMPPlan
{
    class Conversions
    {
        public static double TimeUnit2Todo(string Time, string Unit)
        {
            double dHour;
            string sHour;
            switch (Unit)
            {
                case "UUR":
                    sHour = Time;
                    break;
                case "MIN":
                    dHour = Convert.ToDouble(Time) / 60.0;
                    sHour = dHour.ToString();
                    break;
                default:
                    sHour = "0,0";
                    break;
            }
            return Convert.ToDouble(sHour);
        }
    }
}
