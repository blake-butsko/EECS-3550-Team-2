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
    // The marketing manager selects the plane that should be used for each flight
    internal class MarketingManager : Actor
    {
        string UserId { get; }
        string Password { get; set; }
        public string FName { get; set; }
        public string LName { get; set; }

        string[] PossiblePlanes = { "737", "747", "757", "Norton FalconX 5000" };
        public MarketingManager(string UserId, string Password)
        {
            this.UserId = UserId;
            this.Password = Password;
            populateName();
        }

        private void populateName()
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("EmpList");
            var empTable = worksheet.Tables.Table(0);
            var empIdColumn = empTable.Column(1);

            for (int i = 1; i <= empIdColumn.CellCount(); i++)
            {
                if (string.Equals(UserId, empIdColumn.Cell(i).Value))
                {
                    FName = empIdColumn.Cell(i).CellRight(3).Value.ToString();
                    LName = empIdColumn.Cell(i).CellRight(4).Value.ToString();

                    return;
                }
            }

        }

        public void CreateAccount(string UserId, string Password)
        {
            throw new NotImplementedException();
        }

        // modified flight mmethod (used if we distance isn't a column in active flights) just added parameters
        int CalculateDistances(string Departure, string Arrival)
        {

            var workbook = new XLWorkbook(Globals.databasePath); // Open workbook and worksheet
            var worksheet = workbook.Worksheet("FlightDistance");

            for (int i = 0; i < worksheet.Tables.Count(); i++) // Go through each table in the sheet
            {
                if (String.Equals(Departure, worksheet.Tables.Table(i).Name)) // Get the table that matches the departure location
                {

                    var table = worksheet.Tables.Table(i);

                    for (int j = 1; j <= table.Column(1).CellCount(); j++) // Itterate through all cities in table (Column 1)
                    {

                        if (String.Equals(Arrival, table.Column(1).Cell(j).Value.ToString())) // Get the destination from column
                        {

                            return (int)table.Column(1).Cell(j).CellRight(1).Value; // returns distance


                        }
                    }


                }

            }
            return 0;
        }

        // Function to go into the database and retrieves the flight distance
        // Then assigns a plane dependent on length of flight - To ActiveFlights
        public void ChoosePlane(int FlightId)
        {
            // Code to go into the database and retrieve the flight distance
            // For specified flightId find the distance between Departing and ArrivingAt (add try catch in case of invalid name)
            try
            {
                var workbook = new XLWorkbook(Globals.databasePath); // Open database
                var worksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

                var table = worksheet.Tables.Table(0);

                var idColumn = table.DataRange.Column(1);
                bool IdCheck = false;
                bool IdFound;
                for (int i = 1; i <= idColumn.CellCount(); i++)
                {
                    //Still have to write a try catch only inside this if
                    try { IdFound = string.Equals((string)idColumn.Cell(i).Value, idColumn); }
                    catch { IdFound = false; }
                    
                    if (IdFound)
                    {
                        IdCheck = true;
                        var flightRow = table.DataRange.Row(i);

                        String arrival = (string)flightRow.Cell(2).Value;
                        String departure = (string)flightRow.Cell(3).Value;
                        int distance = CalculateDistances(arrival, departure);
                        String suggested;
                        // code to fetch distance from datasheet thing
                        // Need to find list of planes and distances based on that
                        if (distance < 200)
                        {
                            suggested = PossiblePlanes[0];
                        }
                        else if (distance > 199 && distance < 300)
                            suggested = PossiblePlanes[1];
                        else if (distance > 199 && distance < 300)
                            suggested = PossiblePlanes[2];
                        else
                        {
                            suggested = PossiblePlanes[3];
                        }
                        String userEntry;

                        do
                        {
                            // Row starts at 1 rather than 0
                            Console.WriteLine("This is our suggested plane for this flight {0}", suggested);
                            Console.WriteLine("If this is satisfactory, type: y");
                            Console.WriteLine("If you want to manually enter a plane, type: n");
                            Console.WriteLine("Or if you don't want a plane assigned to this flight, type: quit");
                            // y will set the plane and then quit the function
                            // n will output a series of planes where you put in a number to choose

                            userEntry = Console.ReadLine().ToLower();
                            userEntry = userEntry.Trim();

                            if (String.Equals(userEntry, "y"))
                            {
                                Console.WriteLine("You've selected y, the flight will be updated with the plane");
                                flightRow.Cell(5).Value = suggested;
                                workbook.Save();
                                return;
                            }
                            if (String.Equals(userEntry, "n"))
                            {
                                string userChange;
                                do
                                {
                                    Console.WriteLine("You've selected n, the suggested plane is {0} what would you like to replace it with", suggested);
                                    for (int g = 0; g < PossiblePlanes.Length; g++)
                                    {
                                        Console.WriteLine("{0}. {1}", g, PossiblePlanes[g]);
                                    }
                                    userChange = Console.ReadLine();
                                    userChange = userChange.Trim(); //Might need to remove this
                                    try
                                    {
                                        String planeChoice = PossiblePlanes[Int32.Parse(userChange)];

                                        Console.WriteLine("You've selected {0} is this right type: y/n", planeChoice);
                                        userEntry = Console.ReadLine().ToLower();
                                        userEntry = userEntry.Trim();
                                        if (String.Equals(userEntry, "y"))
                                        {
                                            flightRow.Cell(5).Value = planeChoice;
                                            workbook.Save();
                                            Console.WriteLine("You've selected y, the flight will be updated with the plane");
                                            return;
                                        }

                                        else
                                            Console.WriteLine("Please try again or to leave the program type: quit \n");
                                    }
                                    catch { Console.WriteLine("Invalid input, please try again or to leave type: quit\n"); }
                                } while (!(String.Equals(userChange, "quit")));
                            }
                            Console.WriteLine();
                            if (String.Equals(userEntry, "quit"))
                                return;
                        } while (!(String.Equals(userEntry, "quit")));
                    }
                    else
                    {
                        IdCheck = false;
                    }
                }
                if (!IdCheck)
                {
                    Console.WriteLine("FlightID not found in our database");
                    return;
                }
            }
            catch (FileNotFoundException ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }

            return;
        }

        public void Login(string UserId, string Password)
        {
            throw new NotImplementedException();
        }
    }
}
