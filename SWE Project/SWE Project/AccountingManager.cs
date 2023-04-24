using actor_interface;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    internal class AccountingManager : Actor
    {


        public string UserId;
        string Password;

        public AccountingManager(string UserId, string Password)
        {
            this.UserId = UserId;
            this.Password = Password;
        }

        public void getFlightProfit(string flightId)
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var flightWorksheet = workbook.Worksheet("ActiveFlights");

            var flightTable = flightWorksheet.Tables.Table(0);

            var flightColumn = flightTable.Column(1); //flight id column from flight table

            Flight flight = new Flight();
            bool foundFlight = false;

            // Find flight in database and create object to populate it
            for (int i = 1; i <= flightColumn.CellCount(); i++)
            {
                if (string.Equals((string)flightColumn.Cell(i).Value, flightId))
                {

                    System.DateTime departTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(3).Value.ToString());
                    System.DateTime arrivalTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(4).Value.ToString());

                    flight = new Flight(flightColumn.Cell(i).Value.ToString(),
                        flightColumn.Cell(i).CellRight(1).Value.ToString(),
                         flightColumn.Cell(i).CellRight(2).Value.ToString(),
                         departTime, arrivalTime);


                    foundFlight = true;
                    break;
                }
            }
            if (!foundFlight)
            {
                Console.WriteLine("Flight not found \n");
                return;
            }

            Decimal profit = flight.passengers.Count * flight.Price;

            string fileName = "Flight " + flight.FlightId + "_Manifest.csv";
            FileStream fileCreate = File.Create(fileName);

            string report = "Flight," + (string)flight.FlightId + ",Profit," + profit.ToString();

            using (StreamWriter writer = new StreamWriter(fileCreate))
            {
                writer.Write(report);
            }
        }

        public void getTotalProfit()
        {

            var workbook = new XLWorkbook(Globals.databasePath);
            var flightWorksheet = workbook.Worksheet("ActiveFlights");

            var flightTable = flightWorksheet.Tables.Table(0);

            var flightColumn = flightTable.Column(1); //flight id column from flight table
            List<Flight> flightList = new List<Flight>();
            for (int i = 2; i <= flightColumn.CellCount(); i++)
            {
                Flight flight = new Flight();

                System.DateTime departTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(3).Value.ToString());
                System.DateTime arrivalTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(4).Value.ToString());

                flight = new Flight(flightColumn.Cell(i).Value.ToString(),
                    flightColumn.Cell(i).CellRight(1).Value.ToString(),
                     flightColumn.Cell(i).CellRight(2).Value.ToString(),
                     departTime, arrivalTime);

                flightList.Add(flight);
            }
            string fileName = "Total_profit " + DateTime.Now.ToString() + ".csv";
            fileName = fileName.Replace("/","_");
            fileName = fileName.Replace(":", "_");

            FileStream fileCreate = File.Create(fileName);
            StringBuilder stringBuilder = new StringBuilder();
            stringBuilder.AppendLine("Flight, Profit,");
            Decimal totalProft = 0;

            for (int i = 0; i < flightList.Count; i++)
            {
                Decimal profit = flightList[i].passengers.Count * flightList[i].Price;
                totalProft += profit;
                stringBuilder.Append(flightList[i].FlightId);
                stringBuilder.Append(",");
                stringBuilder.AppendLine(profit.ToString());
            }

            stringBuilder.Append("Total profit,");
            stringBuilder.AppendLine(totalProft.ToString());

            using (StreamWriter writer = new StreamWriter(fileCreate))
            {
                writer.Write(stringBuilder.ToString());
            }

        }

   
    public void Login(string UserId, string Password)
    {
        throw new NotImplementedException();
    }
}
}
