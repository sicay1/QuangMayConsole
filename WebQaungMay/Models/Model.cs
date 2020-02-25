using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace WebQaungMay.Models
{
    public class Model
    {
        public List<K_Year> ky = new List<K_Year>();
        public class MTM
        {
            public MTM()
            {
                sValues = new List<double>();
            }
            public string sName { get; set; }
            public List<double> sValues { get; set; }
            public double sRangeMin { get; set; }
            public double sRangeMax { get; set; }
        }

        public class K_Year
        {
            public K_Year()
            {
                KMon = new List<KOfMonth>();
                KMon.Add(new KOfMonth
                {
                    MonName = "Jan",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(31) {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                    },
                    RndNo_DaysInMon = new List<double>(31) {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                    }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Feb",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(28)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(28)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Mar",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Apr",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(30)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(30)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "May",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Jun",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(30)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(30)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Jul",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Aug",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(30)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(30)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Sep",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Oct",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Nov",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(30)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(30)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                });
                KMon.Add(new KOfMonth
                {
                    MonName = "Dec",
                    K_AvgMon = 0.00,
                    K_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        },
                    RndNo_DaysInMon = new List<double>(31)
                        {
                            0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,
                        }
                });
            }
            public string CityName { get; set; }
            public List<KOfMonth> KMon { get; set; }
            public class KOfMonth
            {
                public string MonName { get; set; }
                public double K_AvgMon { get; set; }
                public List<double> K_DaysInMon { get; set; }
                public List<double> RndNo_DaysInMon { get; set; }
            }
        }
    }
}