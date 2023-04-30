using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Office2013.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net;
using System.Numerics;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    internal class Customer
    {
        String[] west = { "Los Angeles", "Las Vegas", "Atlanta", "Miami", "Cleveland" };
        String[] east = { "Detroit", "New York City", "Salt Lake City", "Washington DC" };
        string[] listOfAirports = { "Nashville", "Cleveland", "Los Angeles", "New York City", "Salt Lake City", "Miami", "Detroit", "Atlanta", "Chicago", "Las Vegas", "Washington DC" };
        string[] listOfAirportsLow = { "nashville", "cleveland", "los angeles", "new york city", "salt lake city", "miami", "detroit", "atlanta", "chicago", "las vegas", "washington dc" };

        public string UserId { get; set; }
        private string Password { get; set; }
        public string FName { get; set; }
        public string LName { get; set; }
        public int Points { get; set; }


        public string CreditCardInfo = "";
        public string wallet = "";
        public string Address = "";
        public int Age = -1;
        public string PhoneNumber = "";
        public Customer(string userId, string password)
        {
            this.UserId = userId;
            this.Password = password;
            populateCustomer();
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

        public void populateCustomer()
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var custWorksheet = workbook.Worksheet("custList");
            var custTable = custWorksheet.Tables.Table(0);

            var custId = custTable.Column(1);

            for (int i = 1; i <= custId.CellCount(); i++)
            {
                if (string.Equals(this.UserId, custId.Cell(i).Value.ToString()))
                {
                    this.FName = custId.Cell(i).CellRight(2).Value.ToString();
                    this.LName = custId.Cell(i).CellRight(3).Value.ToString();
                    this.Address = custId.Cell(i).CellRight(4).Value.ToString();
                    this.PhoneNumber = custId.Cell(i).CellRight(5).Value.ToString();
                    this.Age = Int32.Parse(custId.Cell(i).CellRight(6).Value.ToString());
                    this.wallet = custId.Cell(i).CellRight(7).Value.ToString();
                    this.Points = Int32.Parse(custId.Cell(i).CellRight(8).Value.ToString());
                    this.CreditCardInfo = custId.Cell(i).CellRight(9).Value.ToString();
                }


            }

        }

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
                        var userRow = table.DataRange.Row(i - 1);
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
                                // switch statement to update local variables

                                switch (colIndex)
                                {
                                    case 1:
                                        UserId = newValue;
                                        break;
                                    case 2:
                                        Password = newValue;
                                        break;
                                    case 3:
                                        FName = newValue;
                                        break;
                                    case 4:
                                        LName = newValue;
                                        break;
                                    case 5:
                                        Address = newValue;
                                        break;
                                    case 6:
                                        PhoneNumber = newValue;
                                        break;
                                    case 7:
                                        Age = Int32.Parse(newValue);
                                        break;
                                    case 8:
                                        break;
                                    case 9:
                                        Points = Int32.Parse(newValue);
                                        break;
                                    case 10:
                                        CreditCardInfo = newValue;
                                        break;
                                    default:
                                        Console.WriteLine("Invalid value");
                                        break;
                                }

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


                    if (!(System.DateTime.Now > departTime.AddHours(-24) && System.DateTime.Now < departTime))
                    {
                        Console.WriteLine("It is too early for you to print this\n");

                        return;
                    }


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
                    Console.WriteLine("You are not on this flight\n");
                    return;
                }
            }
        }

        public void updatePoints(int points)
        {
            var workbook = new XLWorkbook(Globals.databasePath); // Open workbook and worksheet
            var worksheet = workbook.Worksheet("CustList");

            var table = worksheet.Tables.Table(0);

            var userId = table.DataRange.Column(1);
            var pointsCol = table.DataRange.Column(9);
            //9 is points
            try
            {
                for (int i = 1; i <= userId.CellCount(); i++)
                {
                    //&& string.Equals((arrival.Cell(i).Value).ToString(), arrivalIn)
                    if (string.Equals((userId.Cell(i).Value).ToString(), UserId))
                    {
                        Points += points;
                        pointsCol.Cell(i).SetValue((Int32.Parse((pointsCol.Cell(i).Value).ToString()) + points).ToString());
                        workbook.Save();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine("Points were not updated"); }
        }
        public string getInfo(int choice)
        {
            var workbook = new XLWorkbook(Globals.databasePath); // Open workbook and worksheet
            var worksheet = workbook.Worksheet("CustList");

            var table = worksheet.Tables.Table(0);

            var userId = table.DataRange.Column(1);
            var choiceValue = table.DataRange.Column(choice);
            //9 is points
            //10 is payment
            try
            {
                for (int i = 1; i <= userId.CellCount(); i++)
                {
                    //&& string.Equals((arrival.Cell(i).Value).ToString(), arrivalIn)
                    if (string.Equals((userId.Cell(i).Value).ToString(), UserId))
                    {
                        return (choiceValue.Cell(i).Value).ToString();
                    }
                }
            }
            catch (Exception ex) { Console.WriteLine("Value not found"); return ""; }
            Console.WriteLine("Value not found");
            return "";
        }

        //This function will book a stored flight in customer history
        public void storeFlight(List<string> flight, int price, int points, string status)
        {
            var workbook = new XLWorkbook(Globals.databasePath); // Open workbook and worksheet
            var worksheet = workbook.Worksheet("CustHistory");

            var table = worksheet.Tables.Table(0);

            var flightId = table.DataRange.Column(1);
            /*try
            {*/
            // Book flight
            Console.WriteLine("Booking flight");
            var listOfData = new ArrayList(); // Making list to feed data into Append data function (IEnumerable)
            listOfData.Add(UserId);
            listOfData.Add(flight[0]);
            listOfData.Add((System.DateTime.Parse(flight[2])).ToUniversalTime().ToString("g")); //WONK - just parse to system.date
            listOfData.Add((System.DateTime.Parse(flight[4])).ToUniversalTime().ToString("g")); //WONK
            listOfData.Add(status);
            listOfData.Add(flight[1]);
            listOfData.Add(flight[3]);
            listOfData.Add(points); // points gotten from price/10 -> this updates the database but it would've put dated it down there
            listOfData.Add(points > 0 ? "Card" : "Points"); // if there's points then you know it was bought with card

            if (points > 0)
            {
                updatePoints(points);
            }

            if (!(table.DataRange.FirstRow().Cell(1).Value.IsBlank))
            {
                table.InsertRowsBelow(1); // Put new flight data into list
            }
            else
            {
                Console.WriteLine("Failed to book flight.\n Check Database");
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
            /*}
            catch (FileNotFoundException ex)
            {
                Console.WriteLine("There was an error");
                return;
            }*/
        }


        public void ScheduleFlight()
        {
            string tempWritein;
            System.DateTime dateIn;
            string departureIn;
            string arrivalIn;
            bool roundtrip;
            do
            {
                try
                {
                    Console.WriteLine("thank you for booking with McDonalds airlines");
                    Console.WriteLine("What day would you like to leave (MM/DD/YYYY)");
                    dateIn = DateTime.ParseExact(Console.ReadLine(), "MM/dd/yyyy", CultureInfo.InvariantCulture); // Just gotta check if this compares all we need is the date
                    Console.WriteLine("Is it a round trip (y/n)");
                    roundtrip = String.Equals(Console.ReadLine().ToLower().Trim(), "y") ? true : false;
                    // Add check here
                    Array.ForEach(listOfAirports, (s) => { Console.WriteLine("|{0} |", s); });
                    Console.WriteLine("Where are you leaving from");
                    tempWritein = Console.ReadLine().ToLower().Trim();
                    departureIn = listOfAirportsLow.Contains(tempWritein) ? listOfAirports[Array.IndexOf(listOfAirportsLow, tempWritein)] : throw new Exception();
                    Console.WriteLine("Where do you wanna go");
                    tempWritein = Console.ReadLine().ToLower().Trim();
                    arrivalIn = listOfAirportsLow.Contains(tempWritein) ? listOfAirports[Array.IndexOf(listOfAirportsLow, tempWritein)] : throw new Exception(); ;
                    break;
                    //could you make exception specific to what faulted? by passing it in?
                }
                catch (Exception) { Console.WriteLine("Invalid input, lets try again"); }
            } while (true);

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


            // If connecting flight is needed swaps arrival with the connecting flight
            // Nashville and Chicago serve as hubs and can fly any where so if you want to fly from Washington DC to Cleveland
            // Washington DC -> Chicago -> Cleveland

            // Moved the arrays to the top
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
            List<List<string>> possibleFlights = new List<List<string>>();
            do
            {
                for (int i = 1; i <= departure.CellCount(); i++)
                {
                    //&& string.Equals((arrival.Cell(i).Value).ToString(), arrivalIn)
                    if (string.Equals((departure.Cell(i).Value).ToString(), departureIn) && string.Equals((arrival.Cell(i).Value).ToString(), arrivalIn))
                    {
                        //if (System.DateTime.Parse(departureTime.Cell(1).Value.ToString()) > dateIn && System.DateTime.Parse(departureTime.Cell(1).Value.ToString()) < dateIn.AddDays(6))
                        if (System.DateTime.Parse(departureTime.Cell(i).Value.ToString()).Date == dateIn.Date)
                        {
                            possibleFlights.Add(new List<string> { flightId.Cell(i).Value.ToString(), departure.Cell(i).Value.ToString(), departureTime.Cell(i).Value.ToString(), arrival.Cell(i).Value.ToString(), arrivalTime.Cell(i).Value.ToString() });
                            //if(possibleFlights.Count>10) { break; }// if it's returning too many
                            // Buy ticket code then if the purchase is successful it stores it in the database

                        }
                    }
                }
                if (possibleFlights.Count == 0 && connecting == "")
                {
                    if (east.Contains(departureIn) && east.Contains(arrivalIn))
                    {
                        connecting = arrivalIn;
                        arrivalIn = "Chicago";
                        Console.WriteLine("East connecting flight book");
                    }
                    else if (west.Contains(departureIn) && west.Contains(arrivalIn))
                    {
                        connecting = arrivalIn;
                        arrivalIn = "Nashville";
                        Console.WriteLine("West connecting flight book");
                    }
                }
                else
                {
                    break;
                }
            } while (true);
            if (possibleFlights.Count == 0)
            {
                Console.WriteLine("There are no flights on this given date");
                return;
            }
            else
            {
                // Define at top
                String userEntry;
                String userChoice;
                List<string> planeChoice;
                do
                {
                    if(connecting != "")
                    {
                        Console.WriteLine("Sorry, but our airline does not currently offer any direct flights between these locations");
                        Console.WriteLine("Let's book the first leg of your connecting flight");
                    }
                    Console.WriteLine("Here are your possible flights please select one by using the corresponding digit:");
                    for (int i = 0; i < possibleFlights.Count; i++)
                    {
                        Console.WriteLine("{0}. Departing from {1} at {2}, Arriving at {3} at {4}", i + 1, possibleFlights[i][1], possibleFlights[i][2], possibleFlights[i][3], possibleFlights[i][4]);
                    }
                    userEntry = Console.ReadLine();
                    userEntry = userEntry.Trim(); //Might need to remove this
                    try
                    {
                        planeChoice = possibleFlights[Int32.Parse(userEntry) - 1];

                        do
                        {
                            Console.WriteLine("You want to select this flight is that correct? y/n");
                            Console.WriteLine("{1} to {2} Leaving {3} arrving {4}", planeChoice[0], planeChoice[1], planeChoice[3], planeChoice[2], planeChoice[4]);
                            userChoice = Console.ReadLine();
                            userChoice = userChoice.Trim();
                            // Save to database
                            if (String.Equals(userChoice, "y"))
                            {
                                decimal payment = new Flight(planeChoice[0], planeChoice[1], planeChoice[3], System.DateTime.Parse(planeChoice[2]), System.DateTime.Parse(planeChoice[4])).Price;
                                do
                                {
                                    if (getInfo(10) == "")
                                    {
                                        Console.WriteLine("It seems you don't have a valid payment method please input it below");
                                        accountInformation();
                                    }
                                    else
                                    {
                                        String temp = getInfo(10);
                                        decimal fink = Points / 100;
                                        if (fink - payment >= 0)
                                        {
                                            do
                                            {
                                                Console.WriteLine("Would you like to use your points on this purchase (y/n)");
                                                userChoice = Console.ReadLine();
                                                userChoice = userChoice.ToLower().Trim();
                                                if (String.Equals(userChoice, "y"))
                                                {
                                                    Console.WriteLine("Points used");
                                                    Points = 100 * (int)(fink - payment);
                                                    Console.WriteLine("You have {0} points left", Points);
                                                    break;
                                                }
                                                else if (String.Equals(userChoice, "n")) { break; }
                                                else if (String.Equals(userChoice, "quit")) { return; }
                                                else { Console.WriteLine("Invalid input please try again or type: quit"); }
                                            } while (!(String.Equals(userChoice, "quit")));
                                        }
                                        else
                                        {
                                            Console.WriteLine("Ticket has been bought with payment method on account");
                                        }
                                        storeFlight(planeChoice, (int)payment, ((int)payment) / 10, "Booked");
                                        if (connecting != "")
                                        {
                                            // Call everything from here 
                                            // Finshing connecting
                                            // Then call return trip with round trip true - with flipped destinations
                                            Console.WriteLine("Let's book the next part of your connecting flight");
                                            ScheduleFlight((System.DateTime.Parse(planeChoice[4])).Date, planeChoice[3], connecting, false);
                                            if (roundtrip)
                                            {
                                                // This is the return trip and the connecting flight will call itself
                                                Console.WriteLine("***********************************************************************************************");
                                                Console.WriteLine("Now lets book your return trip");
                                                Console.WriteLine("{0} to {1}", connecting, planeChoice[1]);
                                                Console.WriteLine("How many days do you plan on staying in your destination?"); // Should add try catch to this
                                                ScheduleFlight((System.DateTime.Parse(planeChoice[4])).Date.AddDays(double.Parse(Console.ReadLine())), connecting, planeChoice[1], false);
                                            }
                                        }
                                        else
                                        {
                                            if (roundtrip)
                                            {
                                                Console.WriteLine("Now lets book your return trip");
                                                Console.WriteLine("How many days do you plan on staying in your destination?");
                                                ScheduleFlight((System.DateTime.Parse(planeChoice[4])).Date.AddDays(double.Parse(Console.ReadLine())), planeChoice[3], planeChoice[1], false);
                                            }
                                        }
                                        return;
                                    }
                                } while (true);
                            }
                            else if (String.Equals(userChoice, "n")) { break; } // Need to test if this
                            else { Console.WriteLine("Invalid input please try again or type: quit"); }
                        } while (!(String.Equals(userChoice, "quit")));
                    }
                    catch
                    {
                        Console.WriteLine("Invalid input please try again or type: quit");
                    }
                } while (!(String.Equals(userEntry, "quit")));
            }

            /*
             //Start of input when you run the function
            //Don't have to worry about hour/minute so you'll either have to drop it or not worry based on how the datetime is set up in c#
             Console.WriteLine("What would you like the new value to be? Enter in the format month/day/year");
             
             */
        }
        public void ScheduleFlight(System.DateTime dateIn, string departureIn, string arrivalIn, bool roundtrip)
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

            List<List<string>> possibleFlights = new List<List<string>>();
            do
            {
                for (int i = 1; i <= departure.CellCount(); i++)
                {
                    //&& string.Equals((arrival.Cell(i).Value).ToString(), arrivalIn)
                    if (string.Equals((departure.Cell(i).Value).ToString(), departureIn) && string.Equals((arrival.Cell(i).Value).ToString(), arrivalIn))
                    {
                        //if (System.DateTime.Parse(departureTime.Cell(1).Value.ToString()) > dateIn && System.DateTime.Parse(departureTime.Cell(1).Value.ToString()) < dateIn.AddDays(6))
                        if (System.DateTime.Parse(departureTime.Cell(i).Value.ToString()).Date == dateIn.Date)
                        {
                            possibleFlights.Add(new List<string> { flightId.Cell(i).Value.ToString(), departure.Cell(i).Value.ToString(), departureTime.Cell(i).Value.ToString(), arrival.Cell(i).Value.ToString(), arrivalTime.Cell(i).Value.ToString() });
                            //if(possibleFlights.Count>10) { break; }// if it's returning too many
                            // Buy ticket code then if the purchase is successful it stores it in the database

                        }
                    }
                }
                if (possibleFlights.Count == 0 && connecting == "")
                {
                    if (east.Contains(departureIn) && east.Contains(arrivalIn))
                    {
                        connecting = arrivalIn;
                        arrivalIn = "Chicago";
                        Console.WriteLine("East connecting flight book");
                    }
                    else if (west.Contains(departureIn) && west.Contains(arrivalIn))
                    {
                        connecting = arrivalIn;
                        arrivalIn = "Nashville";
                        Console.WriteLine("West connecting flight book");
                    }
                }
                else
                {
                    break;
                }
            } while (true);
            if (possibleFlights.Count == 0)
            {
                Console.WriteLine("There are no flights on this given date");
                return;
            }
            else
            {
                // Define at top
                String userEntry;
                String userChoice;
                List<string> planeChoice;
                do
                {
                    if (connecting != "")
                    {
                        Console.WriteLine("Sorry, but our airline does not currently offer any direct flights between these locations");
                        Console.WriteLine("Let's book the first leg of your connecting flight");
                    }
                    Console.WriteLine("Here are your possible flights please select one by using the corresponding digit:");
                    for (int i = 0; i < possibleFlights.Count; i++)
                    {
                        Console.WriteLine("{0}. Departing from {1} at {2}, Arriving at {3} at {4}", i + 1, possibleFlights[i][1], possibleFlights[i][2], possibleFlights[i][3], possibleFlights[i][4]);
                    }
                    userEntry = Console.ReadLine();
                    userEntry = userEntry.Trim(); //Might need to remove this
                    planeChoice = possibleFlights[Int32.Parse(userEntry) - 1];
                    /*try
                    {*/
                    do
                    {
                        Console.WriteLine("You want to select this flight is that correct? y/n");
                        Console.WriteLine("{1} to {2} Leaving {3} arrving {4}", planeChoice[0], planeChoice[1], planeChoice[3], planeChoice[2], planeChoice[4]);
                        userChoice = Console.ReadLine();
                        userChoice = userChoice.Trim();
                        // Save to database
                        if (String.Equals(userChoice, "y"))
                        {
                            decimal payment = new Flight(planeChoice[0], planeChoice[1], planeChoice[3], System.DateTime.Parse(planeChoice[2]), System.DateTime.Parse(planeChoice[4])).Price;
                            do
                            {
                                if (getInfo(10) == "")
                                {
                                    Console.WriteLine("It seems you don't have a valid payment method please input it below");
                                    accountInformation();
                                }
                                else
                                {
                                    String temp = getInfo(10);
                                    decimal fink = Points / 100;
                                    if (fink - payment >= 0)
                                    {
                                        do
                                        {
                                            Console.WriteLine("Would you like to use your points on this purchase (y/n)");
                                            userChoice = Console.ReadLine();
                                            userChoice = userChoice.ToLower().Trim();
                                            if (String.Equals(userChoice, "y"))
                                            {
                                                Console.WriteLine("Points used");
                                                Points = 100 * (int)(fink - payment);
                                                Console.WriteLine("You have {0} points left", Points);
                                                break;
                                            }
                                            else if (String.Equals(userChoice, "n")) { break; }
                                            else if (String.Equals(userChoice, "quit")) { return; }
                                            else { Console.WriteLine("Invalid input please try again or type: quit"); }
                                        } while (!(String.Equals(userChoice, "quit")));
                                    }
                                    else
                                    {
                                        Console.WriteLine("Ticket has been bought with payment method on account");
                                    }
                                    storeFlight(planeChoice, (int)payment, ((int)payment) / 10, "Booked");
                                    if (connecting != "")
                                    {
                                        // Call everything from here 
                                        // Finshing connecting
                                        // Then call return trip with round trip true - with flipped destinations
                                        Console.WriteLine("Let's book the next part of your connecting flight");
                                        ScheduleFlight((System.DateTime.Parse(planeChoice[4])).Date, planeChoice[3], connecting, false);
                                        if (roundtrip)
                                        {
                                            // This is the return trip and the connecting flight will call itself
                                            Console.WriteLine("***********************************************************************************************");
                                            Console.WriteLine("Now lets book your return trip");
                                            Console.WriteLine("How many days do you plan on staying in your destination?");
                                            ScheduleFlight((System.DateTime.Parse(planeChoice[4])).Date.AddDays(double.Parse(Console.ReadLine())), connecting, planeChoice[3], false);
                                        }
                                    }
                                    else
                                    {
                                        if (roundtrip)
                                        {
                                            Console.WriteLine("Now lets book your return trip");
                                            Console.WriteLine("How many days do you plan on staying in your destination?");
                                            ScheduleFlight((System.DateTime.Parse(planeChoice[4])).Date.AddDays(double.Parse(Console.ReadLine())), planeChoice[3], planeChoice[1], false);
                                        }
                                    }
                                    return;
                                }
                            } while (true);
                        }
                        else if (String.Equals(userChoice, "n")) { break; } // Need to test if this
                        else { Console.WriteLine("Invalid input please try again or type: quit"); }
                    } while (!(String.Equals(userChoice, "quit")));
                    //}
                    /*catch
                    {
                        Console.WriteLine("Invalid input please try again or type: quit");
                    }*/
                } while (!(String.Equals(userEntry, "quit")));
            }

            /*
             //Start of input when you run the function
            //Don't have to worry about hour/minute so you'll either have to drop it or not worry based on how the datetime is set up in c#
             Console.WriteLine("What would you like the new value to be? Enter in the format month/day/year");
             
             */
        }

    }

}
