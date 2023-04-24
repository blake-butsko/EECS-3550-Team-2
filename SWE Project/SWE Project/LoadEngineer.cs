using actor_interface;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    // The load engineer is responsible for creating, editing, and deleting flights3
    internal class LoadEngineer : Actor
    {

        public string UserId { get; }
        string Password { get; set; }

        public LoadEngineer(string UserId, string Password)
        {
            this.UserId = UserId;
            this.Password = Password;

        }

        // Create a flight in the database
        public void CreateFlight(string FlightId, string DepartingFrom, string ArrivingAt, System.DateTime DepartTime, System.DateTime ArrivalTime)
        {
            Flight newFlight = new(FlightId, DepartingFrom, ArrivingAt, DepartTime, ArrivalTime);
            try
            {
                var workbook = new XLWorkbook(Globals.databasePath); // Open database
                var worksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

                var table = worksheet.Tables.Table(0); // Get Flight Table

                var listOfData = new ArrayList(); // Making list to feed data into Append data function (IEnumerable)
                listOfData.Add(FlightId);
                listOfData.Add(DepartingFrom);
                listOfData.Add(ArrivingAt);
                listOfData.Add(DepartTime.ToUniversalTime().ToString("g"));
                listOfData.Add(ArrivalTime.ToUniversalTime().ToString("g"));
                if (!(table.DataRange.FirstRow().Cell(1).Value.IsBlank))
                {
                    table.InsertRowsBelow(1); // Put new flight data into list
                }
                else
                {
                    Console.WriteLine("Failed to add flight.\n Check Database");
                    return;
                }

                var tableLastRow = table.LastRow();
                if (listOfData != null)
                {
                    for (int i = 0; i < table.LastRow().CellCount(); i++) // Iterrate through last row of table hitting each cell
                    {

                        tableLastRow.Cell(i + 1).Value = listOfData[i].ToString(); // Change value of cell to list data
                    }

                }
                else
                {
                    Console.WriteLine("Internal Error");
                    return;

                }
                workbook.Save(); // Save changes

            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }



        }
        // Find an existing flight and edit it
        public void EditFlight(string FlightId)
        {
            string[] listOfAirports = { "Nashville", "Cleveland", "Los Angeles", "New York City", "Salt Lake City", "Miami", "Detroit", "Atlanta", "Chicago", "Las Vegas", "Washington DC" };
            // Find flight id in excel file
            try
            {
                var workbook = new XLWorkbook(Globals.databasePath); // Open database
                var flightWorksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

                var flightTable = flightWorksheet.Tables.Table(0);

                var flightIdColumn = flightTable.DataRange.Column(1); // flightId column

                bool updateFlag = false;

                // Look through flight table for flight
                for (int i = 1; i <= flightIdColumn.CellCount(); i++)
                {
                    // Check current flight is in list
                    if (String.Equals(flightIdColumn.Cell(i).Value.GetText(), FlightId.ToString()))
                    {   
                        // Get row of flight
                        var flightRow = flightTable.DataRange.Row(i);

                        String userEntry;

                        do
                        {
                            Console.WriteLine("What field would you like to edit?");
                            Console.WriteLine("To edit flight id, type: flight id");
                            Console.WriteLine("To edit the place the plane is leaving from, type from");
                            Console.WriteLine("To edit the place the plane is arriving, type: to");
                            Console.WriteLine("To edit the date and time the plane is leaving, type: depart-time");
                            Console.WriteLine("To edit the date and time the plane is leaving, type: arrival-time");
                            Console.WriteLine("To stop editing the flight, type: quit");

                            userEntry = Console.ReadLine();
                            // Ensure user unput is valid
                            if (userEntry == null)
                            {
                                Console.WriteLine("Invalid Entry");
                                continue;
                            }

                            userEntry = userEntry.Trim().ToLower();
                            // If the load engineer wants to change the ID
                            if (String.Equals(userEntry, "flight id"))
                            {
                                Console.WriteLine("Current Value: " + flightRow.Cell(1).Value);
                                Console.WriteLine("What would you like the new value to be?");

                                var userChange = Console.ReadLine();

                                if (userChange == null)
                                {
                                    Console.WriteLine("Invalid Entry");
                                    continue;
                                }

                                int tryParseTest;
                                if (!(Int32.TryParse(userChange, out tryParseTest)))
                                {
                                    Console.WriteLine("Invalid input");
                                }
                                else
                                {   
                                    // Check to see if flight id already exists
                                    if (flightIdColumn.Contains(userChange))
                                    {
                                        Console.WriteLine("This flight id already exists.");
                                        return;

                                    }
                                    // Change Id and save
                                    flightRow.Cell(1).Value = userChange;
                                    workbook.Save();
                                    updateFlag = true;
                                }
                            }
                            // If the load engineer wants to change where the flight is departing from
                            else if (String.Equals(userEntry, "from"))
                            {
                                Console.WriteLine("Current Value: " + flightRow.Cell(2).Value);
                                Console.WriteLine("What would you like the new value to be? Enter the city the airport is in: ");

                                string userChange = Console.ReadLine();

                                if (userChange == null)
                                {
                                    Console.WriteLine("Invalid Entry");
                                    continue;
                                }
                                userChange = userChange.ToLower();
                                int tryParseTest;
                                bool validAirport = false;
                                // Check to see if the airport provided is a valid one
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
                                    // Change in database and save
                                    flightRow.Cell(2).Value = userChange;

                                    workbook.Save();
                                    updateFlag = true;
                                }

                            }
                            // If the load engineer wants to change where the flight is arriving at
                            else if (String.Equals(userEntry, "to"))
                            {
                                Console.WriteLine("Current Value: " + flightRow.Cell(3).Value);
                                Console.WriteLine("What would you like the new value to be? Enter the city the airport is in: ");

                                string userChange = Console.ReadLine().ToLower();
                                if (userChange == null)
                                {
                                    Console.WriteLine("Invalid Entry");
                                    continue;
                                }
                                int tryParseTest;

                                bool validAirport = false;
                                // Ensure airport entered is valid
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
                                    // Change in database
                                    flightRow.Cell(3).Value = userChange;
                                    workbook.Save();
                                    updateFlag = true;
                                }

                            }
                            else if (String.Equals(userEntry, "depart-time"))
                            {

                                Console.WriteLine("Current Value: " + flightRow.Cell(4).Value);
                                Console.WriteLine("What would you like the new value to be? Enter in the format month/day/year hour:minute AM/PM");

                                string userChange = Console.ReadLine();
                                if (userChange == null)
                                {
                                    Console.WriteLine("Invalid Entry");
                                    continue;
                                }

                                try
                                {
                                    // Check to see if the date is too far out
                                    System.DateTime newTime = System.DateTime.Parse(userChange);
                                    if(System.DateTime.Now.AddMonths(6) < newTime)
                                    {
                                        Console.WriteLine("This date is too far out.");
                                        continue;
                                    }


                                    flightRow.Cell(4).Value = userChange;
                                    workbook.Save();
                                    updateFlag = true;

                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Invalid Entry");
                                }

                            }
                            else if (String.Equals(userEntry, "arrival-time"))
                            {

                                Console.WriteLine("Current Value: " + flightRow.Cell(5).Value);
                                Console.WriteLine("What would you like the new value to be? Enter in the format month/day/year hour:minute AM/PM");

                                string userChange = Console.ReadLine();
                                if (userChange == null)
                                {
                                    Console.WriteLine("Invalid Entry");
                                    continue;
                                }

                                try
                                {
                                    // Check to see if the date is too far out
                                    System.DateTime newTime = System.DateTime.Parse(userChange);
                                    if (System.DateTime.Now.AddMonths(6) < newTime)
                                    {
                                        Console.WriteLine("This date is too far out to schedule.");
                                        continue;
                                    }


                                    flightRow.Cell(5).Value = userChange;
                                    workbook.Save();
                                    updateFlag = true;

                                }
                                catch (Exception)
                                {
                                    Console.WriteLine("Invalid Entry");
                                }

                            }


                            Console.WriteLine();
                        } while (!(String.Equals(userEntry, "quit")));

                        // If there has been an update in the database make changes in the customer history portion of the database
                        if (updateFlag)
                        {
                            // Make flight object to access passenger list
                            Flight updatedFlight = new Flight(flightRow.Cell(1).Value.ToString(),
                                flightRow.Cell(2).Value.ToString(),
                                flightRow.Cell(3).Value.ToString(),
                                System.DateTime.Parse(flightRow.Cell(4).Value.ToString()),
                                System.DateTime.Parse(flightRow.Cell(5).Value.ToString()));
                            // Get access to customer history sheet and table
                            var custHistWorksheet = workbook.Worksheet("CustHistory");
                            var custHistTable = custHistWorksheet.Tables.Table(0);

                            var custHistIdColumn = custHistTable.Column(1);

                            int passengerIndex = 0;
                            // Update the customer history sheet
                            for (int j = 1; j <= custHistIdColumn.CellCount(); j++)
                            {
                                

                                if (string.Equals(custHistIdColumn.Cell(j).Value.ToString(), updatedFlight.passengers[passengerIndex].UserId)) 
                                {
                                    custHistIdColumn.Cell(j).CellRight(1).Value = updatedFlight.FlightId;
                                    custHistIdColumn.Cell(j).CellRight(2).Value = updatedFlight.departTime.ToString();
                                    custHistIdColumn.Cell(j).CellRight(3).Value = updatedFlight.arrivalTime.ToString();
                                    custHistIdColumn.Cell(j).CellRight(5).Value = updatedFlight.FlightFrom;
                                    custHistIdColumn.Cell(j).CellRight(6).Value = updatedFlight.FlightTo;
                                   
                                    passengerIndex++;
                                }
                               
                            }
                            workbook.Save();
                            return;

                        }

                        else
                            continue;

                    }

                }
                Console.WriteLine("Flight not found \n");
                return;

            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }
        }
        // Find an existing flight and delete it
        public void DeleteFlight(string FlightId)
        {

            // Find flight id in excel file
            try
            {
                var workbook = new XLWorkbook(Globals.databasePath); // Open database
                var flightWorksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

                var flightTable = flightWorksheet.Tables.Table(0);

                var idColumn = flightTable.DataRange.Column(1);
                // Find flight in active flights
                for (int i = 1; i <= idColumn.CellCount(); i++)
                {
                    if (String.Equals(idColumn.Cell(i).Value.ToString(), FlightId))
                    {
                        var flightRow = flightTable.DataRange.Row(i);
                        // Make flight object to get access to passenger list
                        Flight deletedFlight = new Flight(flightRow.Cell(1).Value.ToString(),
                            flightRow.Cell(2).Value.ToString(),
                            flightRow.Cell(3).Value.ToString(),
                            System.DateTime.Parse(flightRow.Cell(4).Value.ToString()),
                             System.DateTime.Parse(flightRow.Cell(5).Value.ToString()));

                        flightRow.Delete();
               
                        // Get Customer History sheet to change flight status
                        var custHistWorksheet = workbook.Worksheet("CustHistory");
                        var custHistTable = custHistWorksheet.Tables.Table(0);
                        var custHistIdColumn = custHistTable.Column(1);
                        int passengerIndex = 0;

                        //Get Customer sheet to refund points
                        
                        var custWorksheet = workbook.Worksheet("CustList");
                        var custTable = custWorksheet.Tables.Table(0);
                        var custIdColumn = custTable.Column(1);
                        List <bool> pointFlags = new List<bool>();
                        // Update the customer history
                        for (int j = 1; j <= custHistIdColumn.CellCount(); j++)
                        {
                            if (string.Equals(custHistIdColumn.Cell(j).Value.ToString(), deletedFlight.passengers[passengerIndex].UserId))
                            {
                                if (string.Equals(custHistIdColumn.Cell(j).CellRight(8).Value.ToString(),"Points"))
                                    pointFlags.Add(true);
                                else
                                    pointFlags.Add(false);

                                
                                custHistIdColumn.Cell(j).CellRight(4).Value = "Canceled";
                                passengerIndex++;
                                if (passengerIndex >= deletedFlight.passengers.Count)
                                    break;

                            }
                           
                        }
                        passengerIndex = 0;
                        for(int j = 1; j <= custIdColumn.CellCount(); j++)
                        {
                            if (string.Equals(custIdColumn.Cell(j).Value.ToString(), deletedFlight.passengers[passengerIndex].UserId))
                            {
                                if (pointFlags[passengerIndex])
                                    custIdColumn.Cell(j).CellRight(8).Value = (int)custIdColumn.Cell(j).CellRight(8).Value + deletedFlight.PointsGenerated;

                                passengerIndex++;
                                if (passengerIndex >= deletedFlight.passengers.Count)
                                    break;
                            }

                        }
                        Console.WriteLine("Flight " + FlightId + " has been deleted.");
                        workbook.Save();


                        return;

                    }
                }
                Console.WriteLine("Flight not found \n");
                return;
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }

        }

        public void Login(string UserId, string Password)
        {
            throw new NotImplementedException();
        }
    }
}

