using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;

namespace StackedBarChartInPPt.Data
{
    public class ChartData
    {
        public SchemeColorValues Color {get; set;}
        public String Name {get; set;}
        public double[] Values{get;set;}
    }
     public sealed class ChartDataCollection
    {
        static List<ChartData> _chartData;

        public static List<ChartData> ChartDataList
        {
            private set {}
            get {
                return _chartData;
            }
        }
        static ChartDataCollection(){
            Initialize();
        }

        private static void Initialize()
        {
            _chartData= new List<ChartData>{
                new() {
                    Color=SchemeColorValues.Accent1,
                    Name="Liza",
                    Values= new double[]{4.3,2.5,3.5,4.5}
                },
                new() {
                    Color=SchemeColorValues.Accent2,
                    Name="Mary",
                    Values=new double[]{2.4,4.4,1.8,2.8}
                },
                  new() {
                    Color=SchemeColorValues.Accent3,
                    Name="Zera",
                    Values=new double[]{2,4,1,3}
                }};
       }
    }

    public struct Months{
        public static string[] Short={
            "July",
            "Aug",
            "Sept",
            "Oct"
        };
    }
}
