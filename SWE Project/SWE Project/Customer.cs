
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using static ClosedXML.Excel.XLPredefinedFormat;

namespace SWE_Project
{
    internal class Customer
    {
        public string UserId { get; }
        private string Password { get; set; }
        public string FName { get; set; }
        public string LName { get; set; }
        public int Points { get; }


        string CreditCardInfo = ""; // Could make into list to hold several cards
        string Email = "";
        string Address = "";
        public int Age = -1;
        string PhoneNumber = "";



        public Customer(string userId, string password, int points, string creditCardInfo, string email, string address, int age, string phoneNumber)
        {
            UserId = userId;
            Password = password;
            Points = points;
            CreditCardInfo = creditCardInfo;
            Email = email;
            Address = address;
            Age = age;
            PhoneNumber = phoneNumber;
        }

        public Customer(string UserId, string FName, string LName, int Age)
        {
            this.UserId = UserId;
            this.FName = FName;
            this.LName = LName;
            this.Age = Age;
        }

        public void custHistory()
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("CustHistory");
            var table = worksheet.Tables.Table(0);
            var flightColumn = table.Column(1); //flight id column
            Console.WriteLine("*********************************************************************************************\n");
            Console.WriteLine("Your Flight History\n");
            Console.WriteLine(flightColumn.Cell(1).CellRight(2).Value.ToString() + " " + flightColumn.Cell(1).CellRight(3).Value.ToString() + " " + flightColumn.Cell(1).CellRight(4).Value.ToString() + " " + flightColumn.Cell(1).CellRight(5).Value.ToString() + " " + flightColumn.Cell(1).CellRight(6).Value.ToString() + " " + flightColumn.Cell(1).CellRight(7).Value.ToString() + " " + flightColumn.Cell(1).CellRight(8).Value.ToString());
            System.DateTime dateTime;
            System.DateTime dateTimeArrive;
            for (int i = 1; i <= flightColumn.CellCount(); i++)
            {
                if (string.Equals((string)flightColumn.Cell(i).Value, UserId))
                {
                    dateTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(3).Value.ToString());
                    dateTimeArrive = System.DateTime.Parse(flightColumn.Cell(i).CellRight(4).Value.ToString());

                    if (string.Equals((string)flightColumn.Cell(i).CellRight(8).Value, "Card"))
                    {
                        Console.WriteLine(flightColumn.Cell(i).CellRight(1).Value.ToString(), flightColumn.Cell(i).CellRight(2).Value.ToString(), dateTime, dateTimeArrive, flightColumn.Cell(i).CellRight(5).Value.ToString(), flightColumn.Cell(i).CellRight(6).Value.ToString(), " $", flightColumn.Cell(i).CellRight(7).Value.ToString(), flightColumn.Cell(i).CellRight(8).Value.ToString());

                    }
                    else
                    {
                        Console.WriteLine(flightColumn.Cell(i).CellRight(1).Value.ToString(), flightColumn.Cell(i).CellRight(2).Value.ToString(), dateTime, dateTimeArrive, flightColumn.Cell(i).CellRight(5).Value.ToString(), flightColumn.Cell(i).CellRight(6).Value.ToString(), " $", flightColumn.Cell(i).CellRight(7).Value.ToString(), flightColumn.Cell(i).CellRight(8).Value.ToString());

                    }
                }
            }
        }

        public void printBoardingPass(string flightId)
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("ActiveFlights");
            var table = worksheet.Tables.Table(0);
            var flightColumn = table.Column(1); //flight id column
            Flight flight = new Flight();
            bool onFlight = false;

            // Find flight in database
            for (int i = 1; i <= flightColumn.CellCount(); i++)
            {
                if (string.Equals(flightId, flightColumn.Cell(i).Value.ToString()))
                {
                    //Get flight from database
                    System.DateTime departTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(3).Value.ToString());
                    System.DateTime arrivalTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(4).Value.ToString());

                    flight = new Flight(flightId, flightColumn.Cell(i).CellRight(1).Value.ToString(), // Debating on keeping in favor of just using rows but keeping for plane type? 
                        flightColumn.Cell(i).CellRight(2).Value.ToString(),
                        departTime, arrivalTime);

                   

                    // Ensure customer is on flight
                    foreach (var customer in flight.passengers)
                    {
                        if (string.Equals(customer.UserId, this.UserId))
                            onFlight = true;
                        
                          
                        
                        if (onFlight)
                        {
                            if (!(System.DateTime.Now > departTime.AddHours(-24) && System.DateTime.Now < departTime))
                            {
                                Console.WriteLine("It is too early for you to print this\n");

                                return;
                            }

                            // Build and print boarding pass
                            Console.WriteLine();
                            StringBuilder builder = new StringBuilder();
                            builder.AppendLine("----------------------------------------------------------------------------------------------");
                            builder.Append("Flight: ");
                            builder.AppendLine(flightId);
                            builder.Append("Passenger: ");
                            builder.Append(this.FName);
                            builder.Append(" ");
                            builder.AppendLine(this.LName);
                            builder.Append("Departing from: ");
                            builder.AppendLine(flight.FlightFrom);
                            builder.Append("Arriving at: ");
                            builder.AppendLine(flight.FlightTo);
                            builder.Append("Departing at: ");
                            builder.AppendLine(flight.departTime.ToString());
                            builder.Append("Arriving at: ");
                            builder.AppendLine(flight.arrivalTime.ToString());
                            builder.AppendLine("----------------------------------------------------------------------------------------------");

                            Console.WriteLine(builder.ToString());
                            return;
                        }
                       
                    }
                    Console.WriteLine("You are not on this flight\n");
                    return;
                }
            }
        }
    }
}
