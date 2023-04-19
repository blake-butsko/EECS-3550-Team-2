using actor_interface;
using ClosedXML.Excel;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    // The load engineer is responsible for creating, editing, and deleting flights
    internal class LoadEngineer : Actor
    {

        string UserId { get; }
        string Password { get; set; }

        public LoadEngineer(string UserId, string Password)
        {
            this.UserId = UserId;
            this.Password = Password;

        }

        public void CreateFlight(int FlightId, Location DepartingFrom, Location ArrivingAt, System.DateTime DateTimeInformation)
        {
            Flight newFlight = new(FlightId, DepartingFrom, ArrivingAt, DateTimeInformation);
            try
            {
                var workbook = new XLWorkbook(Globals.databasePath); // Open database
                var worksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

                var table = worksheet.Tables.Table(0); // Get Flight Table

                var listOfData = new ArrayList(); // Making list to feed data into Append data function (IEnumerable)
                listOfData.Add(FlightId);
                listOfData.Add(DepartingFrom.airport);
                listOfData.Add(ArrivingAt.airport);
                listOfData.Add(DateTimeInformation.ToUniversalTime().ToString("g"));
                if (!(table.DataRange.FirstRow().Cell(1).Value.IsBlank))
                {
                    table.InsertRowsBelow(1); // Put new flight data into list
                }

                var tableLastRow = table.LastRow();
                for (int i = 0; i < table.LastRow().CellCount(); i++) // Iterrate through last row of table hitting each cell
                {

                    tableLastRow.Cell(i + 1).Value = listOfData[i].ToString(); // Change value of cell to list data



                }
                workbook.Save(); // Save changes

            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }



        }

        public void EditFlight(int FlightId)
        {
            string[] listOfAirports = { "Nashville", "Cleveland", "Los Angeles", "New York City", "Salt Lake City", "Miami", "Detroit", "Atlanta", "Chicago", "Las Vegas", "Washington DC" };
            // Find flight id in excel file
            try
            {
                var workbook = new XLWorkbook(Globals.databasePath); // Open database
                var worksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

                var table = worksheet.Tables.Table(0);

                var idColumn = table.DataRange.Column(1);
                for (int i = 1; i <= idColumn.CellCount(); i++)
                {
                    if (String.Equals(idColumn.Cell(i).Value.GetText(), FlightId.ToString()))
                    {
                        var flightRow = table.DataRange.Row(i);

                        String userEntry;

                        do
                        {
                            Console.WriteLine("What field would you like to edit?");
                            Console.WriteLine("To edit flight id, type: flight id");
                            Console.WriteLine("To edit the place the plane is leaving from, type from");
                            Console.WriteLine("To edit the place the plane is arriving, type: to");
                            Console.WriteLine("To edit the date and time the plane is leaving, type: date");
                            Console.WriteLine("To stop editing the flight, type: quit");

                            ; userEntry = Console.ReadLine().ToLower();
                            userEntry = userEntry.Trim();

                            if (String.Equals(userEntry, "flight id"))
                            {
                                Console.WriteLine("Current Value: " + flightRow.Cell(1).Value);
                                Console.WriteLine("What would you like the new value to be?");

                                string userChange = Console.ReadLine();
                                int tryParseTest;
                                if (!(Int32.TryParse(userChange, out tryParseTest)))
                                {
                                    Console.WriteLine("Invalid input");
                                }
                                else
                                {
                                    flightRow.Cell(1).Value = userChange;
                                    workbook.Save();
                                }
                            }
                            else if (String.Equals(userEntry, "from"))
                            {
                                Console.WriteLine("Current Value: " + flightRow.Cell(2).Value);
                                Console.WriteLine("What would you like the new value to be? Enter the city the airport is in: ");

                                string userChange = Console.ReadLine().ToLower();

                                int tryParseTest;
                                bool validAirport = false;
                                for (int j = 0; j < listOfAirports.Length; j++)
                                {
                                    if (userChange.IndexOf(listOfAirports[j].ToLower()) != -1)
                                    {
                                        validAirport = true;
                                    }
                                }
                                if (!(Int32.TryParse(userChange, out tryParseTest) || validAirport))
                                {
                                    Console.WriteLine("Invalid input");
                                }
                                else
                                {
                                    flightRow.Cell(2).Value = userChange;

                                    workbook.Save();
                                }

                            }
                            else if (String.Equals(userEntry, "to"))
                            {
                                Console.WriteLine("Current Value: " + flightRow.Cell(3).Value);
                                Console.WriteLine("What would you like the new value to be? Enter the city the airport is in: ");

                                string userChange = Console.ReadLine().ToLower();

                                int tryParseTest;

                                bool validAirport = false;
                                for (int j = 0; j < listOfAirports.Length; j++)
                                {
                                    if (userChange.IndexOf(listOfAirports[j].ToLower()) != -1)
                                    {
                                        validAirport = true;
                                    }
                                }

                                if (!(Int32.TryParse(userChange, out tryParseTest) || validAirport))
                                {
                                    Console.WriteLine("Invalid input");
                                }
                                else
                                {
                                    flightRow.Cell(3).Value = userChange;
                                    workbook.Save();
                                }

                            }
                            else if (String.Equals(userEntry, "date"))
                            {

                                Console.WriteLine("Current Value: " + flightRow.Cell(4).Value);
                                Console.WriteLine("What would you like the new value to be? Enter in the format month/day/year hour:minute AM/PM");

                                string userChange = Console.ReadLine();

                                try
                                {
                                    System.DateTime newTime = System.DateTime.Parse(userChange);

                                    flightRow.Cell(4).Value = userChange;
                                    workbook.Save();

                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Invalid Entry");
                                }



                            }


                            Console.WriteLine();
                        } while (!(String.Equals(userEntry, "quit")));


                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }



        }

        public void DeleteFlight(int FlightId)
        {

            // Find flight id in excel file
            try
            {
                var workbook = new XLWorkbook(Globals.databasePath); // Open database
                var worksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

                var table = worksheet.Tables.Table(0);

                var idColumn = table.DataRange.Column(1);
                for (int i = 1; i <= idColumn.CellCount(); i++)
                {
                    if (String.Equals(idColumn.Cell(i).Value.GetText(), FlightId.ToString()))
                    {
                        var flightRow = table.DataRange.Row(i);

                        flightRow.Delete();
                        workbook.Save();
                        Console.WriteLine("Flight " + FlightId + " has been deleted.");

                        // Update customer history

                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }


        }

        public void CreateAccount(string UserId, string Password)
        {
            throw new NotImplementedException();
        }


    }
}

