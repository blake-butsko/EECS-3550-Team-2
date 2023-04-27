using actor_interface;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    internal class Customer : Actor
    {
        String[] west = { "Nashville", "Los Angeles", "Las Vegas", "Atlanta", "Miami", "Cleveland" };
        String[] east = { "Chicago", "Detroit", "New York", "Salt Lake", "Washington DC" };

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
        public Customer(string UserId)
        {
            this.UserId = UserId;
        }

        public void CreateAccount(string UserId, string Password)
        {
            throw new NotImplementedException();
        }

        // ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
        public void accountInformation()
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("CustList");
            var table = worksheet.Tables.Table(0);
            var userIdColumn = table.Column(1);
            var header = table.Row(1);

            int colIndex = 0;

            Console.WriteLine("*********************************************************************************************\n");
            Console.WriteLine("Your Account Information\n");

            int userIdRowIndex = -1;
            for (int i = 1; i <= userIdColumn.CellCount(); i++)
            {
                if (string.Equals((string)userIdColumn.Cell(i).Value, UserId))
                {
                    do
                    {
                        var userRow = table.DataRange.Row(i);
                        for (int j = 1; j <= userRow.CellCount(); j++)
                        {
                            Console.WriteLine($"{j}. {header.Cell(j).Value} = {userRow.Cell(j).Value}");
                        }
                        try
                        {
                            Console.WriteLine("Enter the column number for which you'd like to change the value or enter 0 to exit: ");
                            colIndex = int.Parse(Console.ReadLine());
                            if (colIndex == 0)
                                return;

                            Console.WriteLine($"Enter a new value for column {userRow.Cell(colIndex).Value}: ");
                            try
                            {
                                string newValue = Console.ReadLine().Trim();

                                var cellToChange = userRow.Cell(colIndex);
                                cellToChange.SetValue(newValue);
                                workbook.Save();
                                do
                                {
                                    Console.WriteLine("Would you like to modify any other values y/n");
                                    newValue = Console.ReadLine().ToLower().Trim();
                                    if (String.Equals(newValue, "y")) { Console.WriteLine(); break; }
                                    else if (String.Equals(newValue, "n")) { Console.WriteLine("Goodwork!"); return; }
                                    else if (String.Equals(newValue, "quit")) { Console.WriteLine("See you later!"); return; }
                                    else { Console.WriteLine("Invalid input, please try again or type: quit"); }
                                } while (!(String.Equals(newValue, "quit")));
                            }
                            catch (Exception ex) { Console.WriteLine("Invalid input, please try again or type: quit"); }
                        }
                        catch (Exception ex) { Console.WriteLine("Invalid input, please try again or type: quit"); }
                        
                    } while (colIndex != 0); // careful on this
                }
            }

            if (userIdRowIndex != -1)
            {

            }
            else
            {
                Console.WriteLine($"User ID {UserId} not found in database.");
            }
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
                }
            }
        }


        public void Login(string UserId, string Password)
        {
            throw new NotImplementedException();
        }

        //This function will book a stored flight in customer history
        public void storeFlight(string FlightID, System.DateTime departTime, System.DateTime arrivalTime, string status, string depart, string arrival, string points, string payment)
        {
            var workbook = new XLWorkbook(Globals.databasePath); // Open workbook and worksheet
            var worksheet = workbook.Worksheet("CustHistory");

            var table = worksheet.Tables.Table(0);

            var flightId = table.DataRange.Column(1);
            try
            {
                // Book flight

                var listOfData = new ArrayList(); // Making list to feed data into Append data function (IEnumerable)
                listOfData.Add(UserId); // huh
                listOfData.Add(FlightID);
                listOfData.Add(departTime.ToUniversalTime().ToString("g"));
                listOfData.Add(arrivalTime.ToUniversalTime().ToString("g"));
                listOfData.Add(status);
                listOfData.Add(depart);
                listOfData.Add(arrival);
                listOfData.Add(points);
                listOfData.Add(payment);

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


        public void ScheduleFlight(System.DateTime dateIn, string departureIn, string arrivalIn)
        {
            // Write command line part

            // Basic implementation of just checking dates and departure/destination
            // Check inputs against airport list
            // There are none at this date here are the three closest in date that match arrival/departure
            // Have to write code to deal with date and time and how to aprse them - could ask daniel about this
            // 2-44 have no time so start with this 
            var workbook = new XLWorkbook(Globals.databasePath); // Open workbook and worksheet
            var worksheet = workbook.Worksheet("ActiveFlights");

            var table = worksheet.Tables.Table(0);

            var flightId = table.DataRange.Column(1);
            var departure = table.DataRange.Column(2);
            var arrival = table.DataRange.Column(3);
            var departureTime = table.DataRange.Column(4);
            var arrivalTime = table.DataRange.Column(5);

            String connecting = "";
            /*System.DateTime departTime;
            System.DateTime arrivalTime;*/

            String[] west = { "Nashville", "Los Angeles", "Las Vegas", "Atlanta", "Miami", "Cleveland" };
            String[] east = { "Chicago", "Detroit", "New York", "Salt Lake", "Washington DC" };
            // If connecting flight is needed swaps arrival with the connecting flight
            // Nashville and Chicago serve as hubs and can fly any where so if you want to fly from Washington DC to Cleveland
            // Washington DC -> Chicago -> Cleveland
            if (west.Contains(departureIn) && east.Contains(arrivalIn))
            {
                connecting = arrivalIn;
                arrivalIn = "Nashville";
            }
            else if (east.Contains(departureIn) && west.Contains(arrivalIn))
            {
                connecting = arrivalIn;
                arrivalIn = "Chicago";
            }
            // Person selects flight from this list given index they put in and the same y/n thing as marketing manager
            // if statement where if this is empty it says no flights where found given parameters

            List<String> possibleFlights = new List<String>(); //Actually need to store them as a dictionary or object to get name/times too
            for (int i = 1; i <= departure.CellCount(); i++)
            {
                //&& string.Equals((arrival.Cell(i).Value).ToString(), arrivalIn)
                if (string.Equals((departure.Cell(i).Value).ToString(), departureIn))
                {
                    //if (System.DateTime.Parse(departureTime.Cell(1).Value.ToString()) > dateIn && System.DateTime.Parse(departureTime.Cell(1).Value.ToString()) < dateIn.AddDays(6))
                    if (System.DateTime.Parse(departureTime.Cell(i).Value.ToString()).Date == dateIn.Date)
                    {
                        Console.WriteLine("Input {0} equal to {1}", dateIn, System.DateTime.Parse(departureTime.Cell(i).Value.ToString()));
                        Console.WriteLine("This flight is viable {0}", flightId.Cell(i).Value.ToString());
                        possibleFlights.Add(flightId.Cell(i).Value.ToString());
                        //if(possibleFlights.Count>10) { break; }// if it's returning too many


                    }
                }

                // check both departure and arrival
                // code date range/how to work with dates - right now I'm just doing the day of because that will be the easiest (but I'll still have code for multiple flights matching that)
                // add the flight to the customer history - will probably have to access whatever customer ID (could pass this in as a parameter
                // else statement if there's no flights in a range
                // DateTime inputtedDate = DateTime.Parse(Console.ReadLine());
                // Track number of passengers +1 when someone books a ticket -1 if someone cancels (careful about that)
            }
            if (possibleFlights == null)
            {
                Console.WriteLine("There are no flights within 5 days of this given date");
                return;
            }
            else
            {
                // Define at top
                String userEntry;
                String userChoice;
                String planeChoice;
                do
                {
                    Console.WriteLine("Here are your possible flights please select one by using the corresponding digit:");
                    for (int i = 0; i < possibleFlights.Count; i++)
                    {
                        Console.WriteLine("{0}. {1}", i + 1, possibleFlights.ElementAt(i));
                    }
                    userEntry = Console.ReadLine();
                    userEntry = userEntry.Trim(); //Might need to remove this
                    planeChoice = possibleFlights.ElementAt(Int32.Parse(userEntry) - 1);
                    try
                    {
                        do
                        {
                            Console.WriteLine("You want to select this flight is that correct? y/n");
                            /*Console.WriteLine("{1}{2}{3}{4}");*/
                            Console.WriteLine(planeChoice);
                            userChoice = Console.ReadLine();
                            userChoice = userChoice.Trim();
                            // Save to database
                            if (String.Equals(userChoice, "y"))
                            {
                                Console.WriteLine("Flight booked");
                                //Will call book flight functions
                                return;
                            }
                            else if (String.Equals(userChoice, "n")) { break; } // Need to test if this
                            else { Console.WriteLine("Invalid input please try again or type: quit"); }
                        } while (!(String.Equals(userChoice, "quit")));
                    }
                    catch
                    {
                        Console.WriteLine("Invalid input please try again or type: quit");
                    }
                    //Here's where we're printing out the Selections
                } while (!(String.Equals(userEntry, "quit")));
            }
            if (connecting != "")
            {
                Console.WriteLine("Congrats you got a connecting flight");
                // put finished flight scheduling in here or could call the function again adding a true to a default parameter set false
            }

            /*
             //Start of input when you run the function
            //Don't have to worry about hour/minute so you'll either have to drop it or not worry based on how the datetime is set up in c#
             Console.WriteLine("What would you like the new value to be? Enter in the format month/day/year");
             
             */
        }

    }

}
