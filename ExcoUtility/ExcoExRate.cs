using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcoUtility
{
    public class ExcoExRate
    {
        private ExcoExRate() { }

        // return convert to cad rate
        public static double GetToCADRate(ExcoCalendar calendar, string currency)
        {
            if (0 == currency.CompareTo("CA"))
            {
                return 1.0;
            }
            else if (0 == currency.CompareTo("US"))
            {
                return USDtoCAD(calendar);
            }
            else if (0 == currency.CompareTo("CP"))
            {
                return PESOtoCAD(calendar);
            }
            else
            {
                throw new Exception("Invalid currency " + currency);
            }
        }

        // return convert to usd rate
        public static double GetToUSDRate(ExcoCalendar calendar, string currency)
        {
            if (0 == currency.CompareTo("CA"))
            {
                if (USDtoCAD(calendar) > 0.00000001)
                {
                    return 1.0 / USDtoCAD(calendar);
                }
                else
                {
                    return 0.0;
                }
            }
            else if (0 == currency.CompareTo("US"))
            {
                return 1.0;
            }
            else if (0 == currency.CompareTo("CP"))
            {
                if (USDtoCAD(calendar) > 0.00000001)
                {
                    return PESOtoCAD(calendar) / USDtoCAD(calendar);
                }
                else
                {
                    return 0.0;
                }
            }
            else
            {
                throw new Exception("Invalid currency " + currency);
            }
        }

        // return convert to peso rate
        public static double GetToPESORate(ExcoCalendar calendar, string currency)
        {
            if (0 == currency.CompareTo("CA"))
            {
                if (PESOtoCAD(calendar) > 0.00000001)
                {
                    return 1.0 / PESOtoCAD(calendar);
                }
                else
                {
                    return 0.0;
                }
            }
            else if (0 == currency.CompareTo("US"))
            {
                if (PESOtoCAD(calendar) > 0.00000001)
                {
                    return USDtoCAD(calendar) / PESOtoCAD(calendar);
                }
                else
                {
                    return 0.0;
                }
            }
            else if (0 == currency.CompareTo("CP"))
            {
                return 1.0;
            }
            else
            {
                throw new Exception("Invalid currency " + currency);
            }
        }

        // convert rate from 1 USD. to x CAD
        // use calendar year/period
        public static double USDtoCAD(ExcoCalendar calendar)
        {
            int year = calendar.GetCalendarYear();
            int month = calendar.GetCalendarMonth();
            switch (year)
            {
                case 2010:
                case 10:
                case 2011:
                case 11:
                    return 0.0;
                case 2012:
                case 12:
                    switch (month)
                    {
                        case 1:
                            return 1.02;
                        case 2:
                            return 1.0;
                        case 3:
                            return 0.99;
                        case 4:
                            return 1.0;
                        case 5:
                            return 0.99;
                        case 6:
                            return 1.03;
                        case 7:
                            return 1.02;
                        case 8:
                            return 1.0;
                        case 9:
                            return 0.99;
                        case 10:
                            return 0.98;
                        case 11:
                            return 1.0;
                        case 12:
                            return 0.99;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2013:
                case 13:
                    switch (month)
                    {
                        case 1:
                            return 0.99;
                        case 2:
                            return 1.0;
                        case 3:
                            return 1.03;
                        case 4:
                            return 1.02;
                        case 5:
                            return 1.01;
                        case 6:
                            return 1.04;
                        case 7:
                            return 1.05;
                        case 8:
                            return 1.03;
                        case 9:
                            return 1.05;
                        case 10:
                            return 1.03;
                        case 11:
                            return 1.04;
                        case 12:
                            return 1.06;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2014:
                case 14:
                    switch (month)
                    {
                        case 1:
                            return 1.06;
                        case 2:
                            return 1.11;
                        case 3:
                            return 1.11;
                        case 4:
                            return 1.11;
                        case 5:
                            return 1.10;
                        case 6:
                            return 1.08;
                        case 7:
                            return 1.07;
                        case 8:
                            return 1.09;
                        case 9:
                            return 1.09;
                        case 10:
                            return 1.12;
                        case 11:
                            return 1.13;
                        case 12:
                            return 1.14;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2015:
                case 15:
                    switch (month)
                    {
                        case 1:
                            return 1.16;
                        case 2:
                            return 1.27;
                        case 3:
                            return 1.25;
                        case 4:
                            return 1.27;
                        case 5:
                            return 1.21;
                        case 6:
                            return 1.24;
                        case 7:
                            return 1.25; // June
                        case 8:
                            return 1.31; // July
                        case 9:
                            return 1.32;
                        case 10:
                            return 1.335;
                        case 11:
                            return 1.308; // OCT
                        case 12:
                            return 1.335;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2016:
                case 16:
                    switch (month)
                    {
                        case 1:
                            return 1.384;
                        case 2:
                            return 1.401;
                        case 3:
                            return 1.353; //feb
                        case 4:
                            return 1.299; //march
                        case 5:
                            return 1.255;
                        case 6:
                            return 1.311;
                        case 7:
                            return 1.2917; // June
                        case 8:
                            return 1.306; // July
                        case 9:
                            return 1.312;
                        case 10:
                            return 1.312;
                        case 11:
                            return 1.341; // OCT
                        case 12:
                            return 1.343;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2017:
                case 17:
                    switch (month)
                    {
                        case 1:
                            return 1.343; //dec 2016
                        case 2:
                            return 1.301;
                        case 3:
                            return 1.328; //feb
                        case 4:
                            return 1.330; //march
                        case 5:
                            return 1.365;
                        case 6:
                            return 1.35;
                        case 7:
                            return 1.298; // June
                        case 8:
                            return 1.249; // July
                        case 9:
                            return 1.254;
                        case 10:
                            return 1.248;
                        case 11:
                            return 1.248; // OCT
                        case 12:
                            return 1.248;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                default:
                    throw new Exception("Invalid year " + year.ToString());
            }
        }

        // convert rate from 1 PESO to x CAD
        public static double PESOtoCAD(ExcoCalendar calendar)
        {
            int year = calendar.GetCalendarYear();

            int month = calendar.GetCalendarMonth();


            switch (year)
            {
                case 2010:
                case 10:
                case 2011:
                case 11:
                    return 0.0;
                    
                case 2012:
                case 12:
                    switch (month)
                    {
                        case 1:
                            return 0.000525;
                        case 2:
                            return 0.000556;
                        case 3:
                            return 0.000558;
                        case 4:
                            return 0.000557;
                        case 5:
                            return 0.000560;
                        case 6:
                            return 0.000564;
                        case 7:
                            return 0.000571;
                        case 8:
                            return 0.000559;
                        case 9:
                            return 0.000540;
                        case 10:
                            return 0.000546;
                        case 11:
                            return 0.000546;
                        case 12:
                            return 0.000548;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2013:
                case 13:
                    switch (month)
                    {
                        case 1:
                            return 0.000563;
                        case 2:
                            return 0.000562;
                        case 3:
                            return 0.000567;
                        case 4:
                            return 0.000556;
                        case 5:
                            return 0.000551;
                        case 6:
                            return 0.000543;
                        case 7:
                            return 0.000546;
                        case 8:
                            return 0.000542;
                        case 9:
                            return 0.000545;
                        case 10:
                            return 0.000539;
                        case 11:
                            return 0.000552;
                        case 12:
                            return 0.000548;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2014:
                case 14:
                    switch (month)
                    {
                        case 1:
                            return 0.000551;
                        case 2:
                            return 0.000551;
                        case 3:
                            return 0.000541;
                        case 4:
                            return 0.000561;
                        case 5:
                            return 0.000567;
                        case 6:
                            return 0.000572;
                        case 7:
                            return 0.000568;
                        case 8:
                            return 0.000580;
                        case 9:
                            return 0.000567;
                        case 10:
                            return 0.000555;
                        case 11:
                            return 0.000547;
                        case 12:
                            return 0.000515;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2015:
                case 15:
                    switch (month)
                    {
                        case 1:
                            return 0.000486;
                        case 2:
                            return 0.000520;
                        case 3:
                            return 0.000520;
                        case 4:
                            return 0.000489;
                        case 5:
                            return 0.000505;
                        case 6:
                            return 0.000492;//<--- MAY
                        case 7:
                            return 0.00048;//<--- JUNE *check income statement
                        case 8:
                            return 0.000456;
                        case 9:
                            return 0.000432;
                        case 10:
                            return 0.000435;
                        case 11:
                            return 0.000453;
                        case 12:
                            return 0.000424;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2016:
                case 16:
                    switch (month)
                    {
                        case 1:
                            return 0.000436;
                        case 2:
                            return 0.000426; //JAN
                        case 3:
                            return 0.000407;
                        case 4:
                            return 0.000431;
                        case 5:
                            return 0.000440;
                        case 6:
                            return 0.000424;//<--- MAY
                        case 7:
                            return 0.000445;//<--- JUNE *check income statement
                        case 8:
                            return 0.000425;//July
                        case 9:
                            return 0.000444;
                        case 10:
                            return 0.000458;
                        case 11:
                            return 0.000447;
                        case 12:
                            return 0.000437;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2017:
                case 17:
                    switch (month)
                    {
                        case 1:
                            return 0.000448;
                        case 2:
                            return 0.000446; //JAN
                        case 3:
                            return 0.000453;
                        case 4:
                            return 0.000462;
                        case 5:
                            return 0.000463;
                        case 6:
                            return 0.000463;//<--- MAY
                        case 7:
                            return 0.000443;//<--- JUNE *check income statement
                        case 8:
                            return 0.000420;//July
                        case 9:
                            return 0.000424;
                        case 10:
                            return 0.000425;
                        case 11:
                            return 0.000425;
                        case 12:
                            return 0.000425;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                case 2018:
                case 18:
                    switch (month)
                    {
                        case 1:
                            return 0.000446;
                        case 2:
                            return 0.000446; //JAN
                        case 3:
                            return 0.000446;
                        case 4:
                            return 0.000446;
                        case 5:
                            return 0.000446;
                        case 6:
                            return 0.000446;//<--- MAY
                        case 7:
                            return 0.000446;//<--- JUNE *check income statement
                        case 8:
                            return 0.000446;//July
                        case 9:
                            return 0.000446;
                        case 10:
                            return 0.000446;    
                        case 11:
                            return 0.000446;
                        case 12:
                            return 0.000446;
                        default:
                            throw new Exception("Invalid month " + month.ToString());
                    }
                default:
                    throw new Exception("Invalid year " + year.ToString());
            }
        }
    }       
}
