using System;
using CsvHelper;
using System.IO;
using System.Globalization;
using System.Linq;

using ExcelDataReader;
using System.Data;
using System.Data.OleDb;


namespace BMLT2Wayne
{
    public class ZipFile { 
        public DataTable ReadExcel(string fileName, string fileExt)
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = @"provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + fileName + ";Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';"; //for below excel 2007  
            else
                conn = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties='Excel 12.0;HDR=NO';"; //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch { }
            }
            return dtexcel;
        }
    }

    class Program
    {
        public static MeetingOnSchedule[] MoS { get; private set; }
        public static object ExcelDataReader { get; private set; }
        public static DataTableCollection DataTableCollection { get; private set; }

        static void Main(string[] args)
        {
            //MeetingOnSchedule[] MoS;

            using (var stream = File.Open("C:\\Users\\callanm\\Downloads\\zip_code_database.xls", FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader ereader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet eresult = ereader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });
                    DataTableCollection = eresult.Tables;
                    //var erecords = ereader.GetRecords<dynamic>().ToList();
                }
            }
            
       
            using (var streamReader = new StreamReader(@"C:\Users\callanm\Downloads\BMLT.csv"))
            {
                
                using (var csvReader = new CsvReader(streamReader, CultureInfo.InvariantCulture))
                {
                    var records = csvReader.GetRecords<Meeting>().ToList();
                    
                    int numMeetings = records.Count;
                    MeetingOnSchedule[] MoS = new MeetingOnSchedule[numMeetings];
                    int i = 0;
                    foreach (Meeting m in records)
                    {
                        MoS[i] = new MeetingOnSchedule()
                        {
                            CommitteeName = m.CommitteeName,
                            ParentName = m.ParentName,
                            Room = m.Room,
                            WheelChr = m.WheelChr,
                            Day = m.Day,
                            Time = m.Time,
                            Place = m.Place,
                            Address = m.Address,
                            City = m.City,
                            State = m.State,
                            Zip = m.Zip,

                        };
                        //MoS[i].CommitteeName = m.CommitteeName;
                        Console.WriteLine(m.CommitteeName);
                        //string tempstring = new string(m.CommitteeName.ToString);
                    }



                }
            }

        }

            
    }
}
