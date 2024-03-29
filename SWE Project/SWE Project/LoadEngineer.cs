﻿
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    // The load engineer is responsible for creating, editing, and deleting flights
    internal class LoadEngineer
    {

        public string UserId { get; }
        string Password { get; set; }

        public string FName { get; set; }
        public string LName { get; set; }

        public LoadEngineer(string UserId, string Password)
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

            for (int i =1; i <= empIdColumn.CellCount(); i++)
            {
                if (string.Equals(UserId, empIdColumn.Cell(i).Value.ToString()))
                {
                    this.FName = empIdColumn.Cell(i).CellRight(3).Value.ToString();
                    this.LName = empIdColumn.Cell(i).CellRight(4).Value.ToString();

                    return;
                }
            }

        }
        // Create a flight in the database
        public void CreateFlight(string FlightId, string DepartingFrom, string ArrivingAt, System.DateTime DepartTime)
        {
            // Ensure valid location was entered
            string[] listOfAirports = { "Nashville", "Cleveland", "Los Angeles", "New York City", "Salt Lake City", "Miami", "Detroit", "Atlanta", "Chicago", "Las Vegas", "Washington DC" };
            if(!(listOfAirports.Contains(DepartingFrom) && listOfAirports.Contains(ArrivingAt)))
            {
                Console.WriteLine("Invalid Location");
                return;
            }


            try
            {
                var workbook = new XLWorkbook(Globals.databasePath); // Open database
                var activeFlightWorksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

                var flightTable = activeFlightWorksheet.Tables.Table(0); // Get Flight Table

                var flightIdColumn = flightTable.Column(1);
                // Check if the flight already exists
                for (int i = 1; i <= flightIdColumn.CellCount(); i++)
                {
                    if (string.Equals(flightIdColumn.Cell(i).Value.ToString(), FlightId))
                    {
                        Console.WriteLine("Flight Already Exists\n");
                        return;
                    }
                }

                System.DateTime ArrivalTime = System.DateTime.Now;
                // Get flight distance sheet
                var distanceSheet = workbook.Worksheet("FlightDistance");
                bool foundTime = false;
                for (int i = 0; i < distanceSheet.Tables.Count(); i++)
                {
                    if(string.Equals(DepartingFrom.Replace(" ",""), distanceSheet.Tables.Table(i).Name))
                    {
                        var airportTable = distanceSheet.Tables.Table(i);
                        var cityColumn = airportTable.Column(1);
                        for(int j = 1; j <= cityColumn.CellCount(); j++)
                        {
                            if(string.Equals(ArrivingAt, cityColumn.Cell(j).Value.ToString()))
                            {
                                
                                ArrivalTime = DepartTime.AddHours(((double)cityColumn.Cell(j).CellRight(2).Value));
                               
                                foundTime = true;
                                break;
                            }
                        }
                        // Stop looking through database if time is found
                        if (foundTime)
                            break;
                    }

                }



                var listOfData = new ArrayList(); // Making list to feed data into Append data function (IEnumerable)
                listOfData.Add(FlightId);

                listOfData.Add(DepartingFrom);
                listOfData.Add(ArrivingAt);
                listOfData.Add(DepartTime.ToString());
                listOfData.Add(ArrivalTime.ToString());
                listOfData.Add(0); // flight type
                listOfData.Add(0); // passengers

                if (!(flightTable.DataRange.FirstRow().Cell(1).Value.IsBlank))
                {
                    flightTable.InsertRowsBelow(1); // Put new flight data into list
                }
                else
                {
                    Console.WriteLine("Failed to add flight.\n Check Database");
                    return;
                }
               
                var tableLastRow = flightTable.LastRow();
                if (listOfData != null)
                {
                    for (int i = 0; i < flightTable.LastRow().CellCount(); i++) // Iterrate through last row of table hitting each cell
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

                // Get approval from marketing manager
                Console.WriteLine("\nPlease get the approval of a marketing manager to confirm this flight");
                Console.Write("UserId: ");
                string userId = Console.ReadLine();
                Console.Write("Password: ");
                string password = Console.ReadLine();
                // Find entered manager in employee database
                var empWorksheet = workbook.Worksheet("empList");
                var empTable = empWorksheet.Tables.Table(0);
                var empIdColumn = empTable.Column(1);
                MarketingManager marketingManager = new MarketingManager();
                bool foundManager = false;

                //Apply encryption
                byte[] tmpNewHash;
                string SavedPass;
                string checkPass;
                SHA512 shaM = new SHA512Managed();
                var tmpSource = ASCIIEncoding.ASCII.GetBytes(password);//Turns inputted password into bytes
                tmpNewHash = shaM.ComputeHash(tmpSource);//Hashes the bytes
                checkPass = Encoding.UTF8.GetString(tmpNewHash);//turns it back into a string

                // Check for valid marketing manager credentials to select a plane type
                for (int i = 1; i <= empIdColumn.CellCount(); i++)
                {
                   
                    if (string.Equals(empIdColumn.Cell(i).Value.ToString(), userId)) // Don't want to check both user id and password to save comparisions
                    {                    
                            marketingManager = new MarketingManager(empIdColumn.Cell(i).Value.ToString(), password);
                            foundManager = true;
                            break;
                    }
                }
                // Delete flight from database 
                if (!foundManager)
                {
                    flightTable.LastRow().Delete();
                    Console.WriteLine("Invalid Credentials - Canceling flight creation\n");
                }
                else
                {
                    string planeType = marketingManager.ChoosePlane(FlightId, true);
                    if (string.Equals(planeType, ""))
                    {
                        flightTable.LastRow().Delete();
                        Console.WriteLine("Invalid Flight Type - Canceling flight creation\n");
                    }
                    else
                    {
                        flightTable.LastRow().Cell(6).Value = planeType;
                    }
                }
                workbook.Save();

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
                    if (String.Equals(flightIdColumn.Cell(i).Value.ToString(), FlightId.ToString()))
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
                        if(deletedFlight.departTime.Date == System.DateTime.Now.Date)
                            if(deletedFlight.departTime.Subtract(System.DateTime.Now).Hours <= 1) 
                            {
                                Console.WriteLine("Too late to delete flight\n");
                                return;
                            }


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
                                else
                                    custIdColumn.Cell(j).CellRight(8).Value = (int)custIdColumn.Cell(j).CellRight(7).Value + deletedFlight.Price;

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

    }
}

