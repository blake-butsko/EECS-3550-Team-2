using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.EMMA;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    internal class FlightManager
    {
        string UserId;
        string Password;

        public FlightManager(string UserId, string Password)
        {
            this.UserId = UserId;
            this.Password = Password;
        }

        public void getFlightManifest(string FlightId)
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("ActiveFlights");//Might need checks for if it has taken off yet
            var table = worksheet.Tables.Table(0);
            var flightColumn = table.Column(1); //flight id column

            string path = "FlieghtManifest"+ FlightId +".csv";
            Flight flight = new Flight();
            bool foundFlight = false;
            for (int i = 1; i <= flightColumn.CellCount(); i++)
            {
                if (string.Equals((string)flightColumn.Cell(i).Value, FlightId))
                {

                    System.DateTime dateTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(3).Value.ToString());

                    flight = new Flight(flightColumn.Cell(i).Value.ToString(),
                        flightColumn.Cell(i).CellRight(1).Value.ToString(),
                         flightColumn.Cell(i).CellRight(2).Value.ToString(),
                         dateTime);

                    foundFlight = true;
                    break;
                }
            }
            if (!foundFlight)
            {
                Console.WriteLine("Could not find flight");
                return;
            }
            //DateTime localDate = DateTime.Now;
            worksheet = workbook.Worksheet("CustList");//Might need checks for if it has taken off yet
            table = worksheet.Tables.Table(0);
            var custColumn = table.Column(1);

            if (File.Exists(path))
            {
                File.Delete(path);//exits program or deletes old file
            }
            FileStream fs = File.Create(path);
            // Find flight in database and create CSV to populate it
            string custID ="";
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("First Name, Last Name, Customer ID,");
            
            for (int i = 0; i < flight.passengers.Count(); i++)
            {
                for(int y = 1; y <= custColumn.CellCount(); y++)
                {
                    if (string.Equals((string)custColumn.Cell(y).Value, flight.passengers.ElementAt(i)))
                    {
                        stringBuilder.Append(custColumn.Cell(y).CellRight(2).Value.ToString());
                        stringBuilder.Append(",");
                        stringBuilder.Append(custColumn.Cell(y).CellRight(3).Value.ToString());
                        stringBuilder.Append(",");
                        stringBuilder.AppendLine(custColumn.Cell(y).Value.ToString());
                        
                    }
                }

            }
            using (StreamWriter writer = new StreamWriter(fs))
            {
                writer.Write(stringBuilder.ToString());
            }
            fs.Close();
            Console.WriteLine("File " + path + " has been created.");
        }
    }

}
