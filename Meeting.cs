using CsvHelper.Configuration.Attributes;
//using Microsoft.Office.Interop.Excel;
//using _Excel = Microsoft.Office.Interop.Excel;
//using Range = Microsoft.Office.Interop.Excel.Range;
//using ExcelDataReader;
using Microsoft.Office.Interop.Excel;

namespace BMLT2Wayne
{
    public class Meeting
    {

        [Name("Committee")]
        public string Committee { get; set; }
        [Name("CommitteeName")]
        public string CommitteeName { get; set; }
        [Name("AddDate")]
        public string AddDate { get; set; }
        [Name("AreaRegion")]
        public string AreaRegion { get; set; }
        [Name("ParentName")]
        public string ParentName { get; set; }
        [Name("ComemID")]
        public string ComemID { get; set; }
        [Name("ContactID")]
        public string ContactID { get; set; }
        [Name("ContactName")]
        public string ContactName { get; set; }
        [Name("CompanyName")]
        public string CompanyName { get; set; }
        [Name("ContactAddrID")]
        public string ContactAddrID { get; set; }
        [Name("ContactAddress1")]
        public string ContactAddress1 { get; set; }
        [Name("ContactAddress2")]
        public string ContactAddress2 { get; set; }
        [Name("ContactCity")]
        public string ContactCity { get; set; }
        [Name("ContactState")]
        public string ContactState { get; set; }
        [Name("ContactZip")]
        public string ContactZip { get; set; }
        [Name("ContactCountry")]
        public string ContactCountry { get; set; }
        [Name("ContactPhone")]
        public string ContactPhone { get; set; }
        [Name("MeetingID")]
        public string MeetingID { get; set; }
        [Name("Room")]
        public string Room { get; set; }
        [Name("Closed")]
        public string Closed { get; set; }
        [Name("WheelChr")]
        public string WheelChr { get; set; }
        [Name("Day")]
        public string Day { get; set; }
        [Name("Time")]
        public string Time { get; set; }
        [Name("Language1")]
        public string Language1 { get; set; }
        [Name("Language2")]
        public string Language2 { get; set; }
        [Name("Language3")]
        public string Language3 { get; set; }
        [Name("LocationId")]
        public string LocationId { get; set; }
        [Name("Place")]
        public string Place { get; set; }
        [Name("Address")]
        public string Address { get; set; }
        [Name("City")]
        public string City { get; set; }
        [Name("LocBorough")]
        public string LocBorough { get; set; }
        [Name("State")]
        public string State { get; set; }
        [Name("Zip")]
        public string Zip { get; set; }
        [Name("Country")]
        public string Country { get; set; }
        [Name("Directions")]
        public string Directions { get; set; }
        [Name("Institutional")]
        public string Institutional { get; set; }
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
        [Name("Delete")]
        public string Delete { get; set; }
        [Name("LastChanged")]
        public string LastChanged { get; set; }
        [Name("Longitude")]
        public string Longitude { get; set; }
        [Name("Latitude")]
        public string Latitude { get; set; }
        [Name("ContactGP")]
        public string ContactGP { get; set; }
        [Name("PhoneMeetingNumber")]
        public string PhoneMeetingNumber { get; set; }
        [Name("VirtualMeetingLink")]
        public string VirtualMeetingLink { get; set; }
        [Name("VirtualMeetingInfo")]
        public string VirtualMeetingInfo { get; set; }
        [Name("TimeZone")]
        public string TimeZone { get; set; }
        [Name("bmlt_id")]
        public string bmlt_id { get; set; }
        [Name("unpublished")]
        public string unpublished { get; set; }

    }
}
