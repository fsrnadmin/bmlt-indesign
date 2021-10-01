using CsvHelper.Configuration.Attributes;
//using Microsoft.Office.Interop.Excel;
//using _Excel = Microsoft.Office.Interop.Excel;
//using Range = Microsoft.Office.Interop.Excel.Range;
//using ExcelDataReader;
using Microsoft.Office.Interop.Excel;

namespace BMLT2Wayne
{
    public class MeetingOnSchedule
    {
        [Name("CommitteeName")]
        public string CommitteeName { get; set; }
        [Name("ParentName")]
        public string ParentName { get; set; }
        [Name("Room")]
        public string Room { get; set; }
        //[Name("Closed")]
       // public string Closed { get; set; }
        [Name("WheelChr")]
        public string WheelChr { get; set; }
        [Name("Day")]
        public string Day { get; set; }
        
        [Name("Time")]
        public string Time { get; set; }
        
        [Name("Place")]
        public string Place { get; set; }
        
        [Name("Address")]
        public string Address { get; set; }
        
        [Name("City")]
        public string City { get; set; }
        
        //[Name("LocBorough")]
        //public string LocBorough { get; set; }
        
        [Name("State")]
        public string State { get; set; }
        
        [Name("Zip")]
        public string Zip { get; set; }
        
        [Name("County")]
        public string County { get; set; }
        
        [Name("Directions")]
        public string Directions { get; set; }
                
        [Name("Format1")]
        public string Format1 { get; set; }
        
        [Name("Format2")]
        public string Format2 { get; set; }
        
        [Name("Format3")]
        public string Format3 { get; set; }
        
        [Name("Format4")]
        public string Format4 { get; set; }
        
        [Name("Format5")]
        public string Format5 { get; set; }
        
        [Name("VirtualMeetingLink")]
        public string VirtualMeetingLink { get; set; }
        
        [Name("VirtualMeetingInfo")]
        public string VirtualMeetingInfo { get; set; }
     
        [Name("unpublished")]
        public string unpublished { get; set; }

    }
}
