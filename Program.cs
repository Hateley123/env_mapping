using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using Impinj.OctaneSdk; // Calls for the reader SDK so I can control the reader requiers SDK to be installed
using OfficeOpenXml; // Requires EPP to be installed (Nuget Packages)

namespace RFID_tag
{  
    class Program
    {
        
        public static class SolutionConstants
        {
            public const string ReaderHostName = "169.254.99.252"; //Reader Ip Address
            // Place to store data while program is running. 
            public const string filePath1 = "C:\\Users\\jmhat\\UCLA\\test_data"; // Place to store data
            public const double tagRadarCrossSection = 0.07*0.017;
            public const double lightSpeed = 299792458;
            public const int antennaGainInDecibals = 11;
            public const double antennaGain = 10^(antennaGainInDecibals/10);
            public const double TxPower = 30;

        }
        /******************************************************************************
         * The following class was created to declare and retrieve variables from various
         * methods used without the code. Counter is being used to determine the number
         * of data sets I have and system data is being used to retrieve and store
         * the tag data
         ******************************************************************************/
        static class VariableDeclarations
        {
            public static int counter;
            public static List<Tag> systemData = new();
            public static List<double> distance = new();
            public static double PowerMultiplier;
            

        }
       
        // Create a new reader object so I am able to communicate with the reader and get all information
        static ImpinjReader reader = new ImpinjReader();
        static bool reader_connected = false;
        static void Main(string[] args)
        {            
            // Initialize counters for data storage
            VariableDeclarations.counter = 1;
            
            double num = (SolutionConstants.antennaGain * SolutionConstants.lightSpeed);
            double denum = Math.Pow(4 * Math.PI, 3);

            VariableDeclarations.PowerMultiplier = Math.Pow(num / denum, 0.25);
            // Opens the excelPackage so I can store the data on an excel spreadsheet for post processing
            ExcelPackage excelPackage = new ExcelPackage();
            ExcelWorksheet xlWorksheet = excelPackage.Workbook.Worksheets.Add("Sheet 1");
            try
            {
                // connects to the reader
                Console.WriteLine("Connecting to reader");       
                reader.Connect("169.254.99.252");
                reader_connected = true;
                
            }
            catch (OctaneSdkException e2)
            {
                try
                {
                    reader.Connect("169.254.1.1");
                    reader_connected = true;

                }
                catch(OctaneSdkException e1)
                {
                    Console.WriteLine(e1.Message);
                }
                
            }

            if(reader_connected)
            { 
                Console.WriteLine("Connection Established");
                // Gets current Settings of the reader
                Settings settings = reader.QueryDefaultSettings();
                /************************************************************************
                 * Section where I describe the reader what to data to report and store
                 * Below when I say IncludeAntennaPortNumber, I will be assigning a value 
                 * for the antenna port for the tag being included. However, since I only
                 * have one antenna right now, I don't really need this statement. But in 
                 * future I will have multiple antennas attached to my drone, so this antenna
                 * number will be important for later. 
                 * 
                 * Right now, I am telling the reader to collect the antenna port number,
                 * phase and peak RSSI
                 ************************************************************************/
                settings.Report.IncludeAntennaPortNumber = true;
                settings.Report.IncludePhaseAngle = true;
                settings.Report.IncludePeakRssi = true;
                settings.Report.IncludeDopplerFrequency = true;
                settings.Report.IncludeChannel = true;
                settings.Report.IncludeFirstSeenTime = true;
                /*settings.Report.IncludeChannel = true;
                settings.Report.IncludeDopplerFrequency = true;*/
                // Reports each time a tag is detected and the values associated with the tag
                settings.Report.Mode = ReportMode.Individual;
                /*****************************************************************************
                 * For RSSI and Angle of Arrival measurments we need to know the frequency the
                 * reader operates at. There are possible ways to get around it but it doesnt 
                 * hurt to set the frequency and it shows how to change the systems operating range.
                 * However, right now I am unable to change the transmission frequency of the system.
                 * So that will be a problem for another day as I can use AOA from multiple anchors to localize the system
                 *****************************************************************************/
                              
                reader.ApplySettings(settings);   
                // Collects tags readings and stores the data into list
                reader.TagsReported += OnTagsReported;
                reader.Start();
                // Collects data until I press enter
                Console.WriteLine("Press Enter to exit");
                Console.ReadLine();      
                reader.Stop(); 
                reader.Disconnect(); 
                Console.WriteLine("Disconnected");
                // Now the data is placed into an excel file
                int row = 1;
                double r;
                double rnum;
                
                foreach (Tag tag in VariableDeclarations.systemData)
                {
                    
                    xlWorksheet.Cells[row, 1].Value = tag.Epc;
                    xlWorksheet.Cells[row, 2].Value = tag.AntennaPortNumber;
                    xlWorksheet.Cells[row, 3].Value = tag.PhaseAngleInRadians;
                    xlWorksheet.Cells[row, 4].Value = tag.PeakRssiInDbm;
                    xlWorksheet.Cells[row, 5].Value = tag.ChannelInMhz;
                    xlWorksheet.Cells[row, 6].Value = tag.FirstSeenTime;
                    rnum = (VariableDeclarations.PowerMultiplier / (tag.ChannelInMhz * 1000000));
                    r =  rnum* Math.Pow(10, (SolutionConstants.TxPower - tag.PeakRssiInDbm) / 20);
                    xlWorksheet.Cells[row, 7].Value = r;
                        
                    row++;
                }
                FileInfo excelFile = new FileInfo("C:\\Users\\jmhat\\UCLA\\test_data\\multi_tag\\test.xlsx");
                excelPackage.SaveAs(excelFile);                                             
                
            }
            else
            {
                Console.WriteLine("Well you done messed up you idiot");

            }
                        
        }
       
        static void OnTagsReported(ImpinjReader sender, TagReport report)
        {
            
            foreach (Tag tag in report)
            {                
                VariableDeclarations.systemData.Add(tag); 
                Console.WriteLine(tag.Epc);
                Console.WriteLine(tag.Epc.ToString());
                
                Console.WriteLine(tag.Epc.ToString().Substring(0,4));


            }
        }
        
              
    }
}
