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
    // The flight manager is responisble for printing the boarding pass for flights 24 hours before they take off
    internal class FlightManager
    {
        public string UserId { get; }
        string Password;
        public string FName { get; set; }
        public string LName { get; set; }
        public FlightManager(string UserId, string Password)
        {
            this.UserId = UserId;
            this.Password = Password;
            populateName();
        }
        // Grab name from database given userId and password
        private void populateName()
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("EmpList");
            var empTable = worksheet.Tables.Table(0);
            var empIdColumn = empTable.Column(1);

            for (int i = 1; i <= empIdColumn.CellCount(); i++)
            {
                if (string.Equals(UserId, empIdColumn.Cell(i).Value.ToString()))
                {
                    this.FName = empIdColumn.Cell(i).CellRight(3).Value.ToString();
                    this.LName = empIdColumn.Cell(i).CellRight(4).Value.ToString();

                    return;
                }
            }

        }

        public void getFlightManifest(string FlightId)
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("ActiveFlights");//Might need checks for if it has taken off yet
            var table = worksheet.Tables.Table(0);
            var flightColumn = table.Column(1); //flight id column

            string path = "FlightManifest"+ FlightId +".csv";
            Flight flight = new Flight();
            bool foundFlight = false;
            for (int i = 1; i <= flightColumn.CellCount(); i++)
            {
                if (string.Equals(flightColumn.Cell(i).Value.ToString(), FlightId))
                {

                    System.DateTime from = System.DateTime.Parse(flightColumn.Cell(i).CellRight(3).Value.ToString());
                    System.DateTime dest = System.DateTime.Parse(flightColumn.Cell(i).CellRight(4).Value.ToString());
                    flight = new Flight(flightColumn.Cell(i).Value.ToString(),
                        flightColumn.Cell(i).CellRight(1).Value.ToString(),
                         flightColumn.Cell(i).CellRight(2).Value.ToString(),
                         from, dest);

                    foundFlight = true;
                    table.Row(i).Delete();
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
            // Build string to write to csv
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
            // Write string to file
            using (StreamWriter writer = new StreamWriter(fs))
            {
                writer.Write(stringBuilder.ToString());
            }
            fs.Close();
            Console.WriteLine("File " + path + " has been created.");
        }
    }

}
