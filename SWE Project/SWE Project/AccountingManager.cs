using actor_interface;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    internal class AccountingManager : Actor
    {


        string UserId;
        string Password;

        public AccountingManager(string UserId, string Password)
        {
            this.UserId = UserId;
            this.Password = Password;
        }

        public void getFlightProfit(string flightId)
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("ActiveFlights");

            var table = worksheet.Tables.Table(0);

            var flightColumn = table.Column(1); //flight id column

            Flight flight = new Flight();
            bool foundFlight = false;

            // Find flight in database and create object to populate it
            for (int i = 1; i <= flightColumn.CellCount(); i++)
            {
                if (string.Equals((string)flightColumn.Cell(i).Value, flightId))
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
            if (!foundFlight && flight == null)
            {
                Console.WriteLine("Flight not found \n");
                return;
            }

            Decimal profit = flight.passengers.Count * flight.Price;

            Console.WriteLine("The profit for this flight is: " + profit); // Replace with csv print out
        }


        public void Login(string UserId, string Password)
        {
            throw new NotImplementedException();
        }
    }
}
