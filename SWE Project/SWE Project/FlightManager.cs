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
        public string UserId { get; }//The Flight managers ID
        string Password;
        public string FName { get; set; }//Flight managers first name
        public string LName { get; set; }//Flight manangers last name
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
                if (string.Equals(UserId, empIdColumn.Cell(i).Value.ToString()))//If the Users is in the Employees list database then get First and last name
                {
                    this.FName = empIdColumn.Cell(i).CellRight(3).Value.ToString();
                    this.LName = empIdColumn.Cell(i).CellRight(4).Value.ToString();

                    return;
                }
            }

        }

        public void getFlightManifest(string FlightId)//Creates a csv for the flightid given if it is going to take off in 24 hours
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("ActiveFlights");//Might need checks for if it has taken off yet
            var table = worksheet.Tables.Table(0);
            var flightColumn = table.Column(1); //flight id column

            string path = "FlightManifest"+ FlightId +".csv";//Create name of csv file to be created
            Flight flight = new Flight();
            bool foundFlight = false;
            var totalRows = worksheet.LastRowUsed().RowNumber();//Gets number of last row filled with info
            for (int i = 1; i <= totalRows; i++)
            {
                if (string.Equals(flightColumn.Cell(i).Value.ToString(), FlightId))//If the flight ID is in the ActiveFlight database it creates a flight object
                {

                    System.DateTime from = System.DateTime.Parse(flightColumn.Cell(i).CellRight(3).Value.ToString());
                    System.DateTime dest = System.DateTime.Parse(flightColumn.Cell(i).CellRight(4).Value.ToString());
                    flight = new Flight(flightColumn.Cell(i).Value.ToString(),
                        flightColumn.Cell(i).CellRight(1).Value.ToString(),
                         flightColumn.Cell(i).CellRight(2).Value.ToString(),
                         from, dest);//Creates flight object

                    foundFlight = true;
                    table.Row(i).Delete();//Removes flight object from sheet doesn't work unless it is saved <----------
                    break;
                }
            }
            if (!foundFlight)//If no flight object was created then the flight ID wasn't in the activeflight database
            {
                Console.WriteLine("Could not find flight\n");
                return;
            }
            //DateTime localDate = DateTime.Now;


            if (File.Exists(path))
            {
                File.Delete(path);//deletes old file
            }
            FileStream fs = File.Create(path);//Create csv to populate

            string custID ="";
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("First Name, Last Name, Customer ID,");//Add Header row to string builder
            
            for (int i = 0; i < flight.passengers.Count(); i++)//For all passengers on the list in flight object add First, last and ID to stringbuilder
            {
                stringBuilder.Append(flight.passengers.ElementAt(i).FName);
                stringBuilder.Append(",");
                stringBuilder.Append(flight.passengers.ElementAt(i).LName);
                stringBuilder.Append(",");
                stringBuilder.AppendLine(flight.passengers.ElementAt(i).UserId);

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
