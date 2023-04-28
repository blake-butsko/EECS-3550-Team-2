﻿/* Authors: Daniel Viar, Issac Wamer, Blake Butsko
 * 
 * Professor: Dr. Thomas
 * 
 * Class: EECS 3550
 * 
 * Airline project - This project is focused on creating an airline experience similar to commercial airlines with accounts, flight scheduling, and rewards
 */

using System.Runtime.CompilerServices;
using ClosedXML;
using ClosedXML.Excel;
using System.Collections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.VariantTypes;
using static ClosedXML.Excel.XLPredefinedFormat;
using SWE_Project;
using System.Text;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Runtime.Intrinsics.Arm;
using DocumentFormat.OpenXml.Office2010.Word;
using System.Net;

// Class for global variables following c# standards
public class Globals
{
    public static string databasePath = "";
}


internal class CLICaller
{
    char state;
    public CLICaller() { } // Constructor

    public void CustomerCli(SWE_Project.Customer person) // add customer object here
    {
        Console.WriteLine("*********************************************************************************************");
        Console.WriteLine("Welcome Back " + person.FName + " " + person.LName + "!\n");
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To book a flight, enter book.");
            Console.WriteLine("To print a boarding pass, enter print.");
            Console.WriteLine("To look at account, enter account.");
            Console.WriteLine("To exit the customer portal, enter quit.\n");


            userInput = Console.ReadLine();
            if (userInput != null)
                userInput = userInput.ToLower();
            else
                Console.WriteLine("Invalid Entry\n");

            Console.WriteLine("");

            if (string.Equals(userInput, "book"))
            {
                // Booking method here
            }
            else if (string.Equals(userInput, "print"))
            {
                Console.WriteLine("Please enter the flight id of the plane you are trying to board: ");
                string userResponse = Console.ReadLine();
                if (userResponse != null)
                    person.printBoardingPass(userResponse);
                
            }
            else if (string.Equals(userInput, "account"))
            {
                person.custHistory();
            }
            else if(!string.Equals(userInput, "quit"))
            {
                Console.WriteLine("*********************************************************************************************");
                System.Environment.Exit(1);
            }
            else
                Console.WriteLine("Invalid Entry\n");
            Console.WriteLine("*********************************************************************************************");

        } while (string.Equals(userInput, "quit"));

        Console.WriteLine("*********************************************************************************************");
        System.Environment.Exit(1);
        return;
    }

    public void LoadEngineerCli(SWE_Project.LoadEngineer engineer)
    {
        Console.WriteLine("*********************************************************************************************");
        string user = engineer.UserId; // Temp
        Console.WriteLine("\n Welcome Back " + engineer.FName + "!\n");
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To create a flight, enter create.");
            Console.WriteLine("To edit a flight, enter edit.");
            Console.WriteLine("To delete a flight, enter delete.");
            Console.WriteLine("To exit the load engineer portal, enter quit.\n");


            userInput = Console.ReadLine();
            if (userInput != null)
                userInput = userInput.ToLower();
            else
                Console.WriteLine("Invalid Entry\n");

            Console.WriteLine("");

            if (string.Equals(userInput, "create"))
            {
                // Booking method here
                Console.Write("Enter an ID for the flight: ");
                string FlightId = Console.ReadLine();
                Console.Write("Enter the airport the flight is taking off from: ");
                string DepartingFrom = Console.ReadLine();
                Console.Write("Enter the airport the flight will be arriving at: ");
                string ArrivingAt = Console.ReadLine();
                Console.Write("Enter the date and time of departure in the format MM/DD/YYYY HH:MM ");
                string DepartTime = Console.ReadLine();

                string confIn;
                do
                {
                    Console.Write("Enter Yes or No (Y/N) to confirm submission: ");
                    confIn = Console.ReadLine();
                    if (confIn == "Y")
                    {
                        try
                        {
                            engineer.CreateFlight(FlightId, DepartingFrom, ArrivingAt, System.DateTime.Parse(DepartTime));
                        }
                        catch
                        {
                            Console.WriteLine("Invalid Time entered");
                        }
                    }
                } while (confIn == "y" || confIn == "n");
            }
            else if (string.Equals(userInput, "edit"))
            {
                Console.Write("Enter the ID for the flight you want to edit: ");
                string FlightId = Console.ReadLine();
                if (FlightId != null)
                {
                    engineer.EditFlight(FlightId);
                }
                else
                {
                    Console.WriteLine("Invalid Entry\n");
                }
            }
            else if (string.Equals(userInput, "delete"))
            {
                Console.Write("Enter the ID for the flight you want to delete: ");
                string FlightId = Console.ReadLine();
                if (FlightId != null)
                {
                    engineer.DeleteFlight(FlightId);
                }
                else
                {
                    Console.WriteLine("Invalid Entry\n");
                }
            }
            else if (string.Equals(userInput, "account"))
            {

            }
            else if (string.Equals(userInput, "quit"))
            {
                Console.WriteLine("*********************************************************************************************");
                System.Environment.Exit(1);
            }
            else
                Console.WriteLine("Invalid Entry\n");
            Console.WriteLine("*********************************************************************************************");

        } while (!string.Equals(userInput, "quit"));

        return;

    }

    public void marketingManagerCli(SWE_Project.MarketingManager marketing)
    {
        Console.WriteLine("*********************************************************************************************");
        Console.WriteLine("Welcome Back " + marketing.FName + " " + marketing.LName + "!\n");
        //string user = marketing.UserId; // Temp
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To select a plane type for a flight, enter plane.");
            Console.WriteLine("To exit the marketing manager portal, enter quit.\n");


            userInput = Console.ReadLine();
            if (userInput != null)
                userInput = userInput.ToLower();
            else
                Console.WriteLine("Invalid Entry\n");

            Console.WriteLine("");

            if (string.Equals(userInput, "plane"))
            {

                Console.WriteLine("Please Eneter what flight you would like to edit/ add");
                string flightId = Console.ReadLine();
                if(flightId != null)
                    marketing.ChoosePlane(flightId,false);

            }
            else if (string.Equals(userInput, "quit"))
            {
                Console.WriteLine("*********************************************************************************************");
                System.Environment.Exit(1);
            }
            else
                Console.WriteLine("Invalid Entry\n");
            Console.WriteLine("*********************************************************************************************");

        } while (!string.Equals(userInput, "quit"));


        Console.WriteLine("*********************************************************************************************");
        System.Environment.Exit(1);
        return;


    }

    public void FlightManagerCli(SWE_Project.FlightManager flighter)
    {
        Console.WriteLine("*********************************************************************************************");
        Console.WriteLine("Welcome Back " + flighter.FName + " " + flighter.LName + "!\n");
        string user = flighter.UserId; // Temp
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To print a flight manifest for a flight, enter print.");
            Console.WriteLine("To exit the marketing manager portal, enter quit.\n");


            userInput = Console.ReadLine();
            if (userInput != null)
                userInput = userInput.ToLower();
            else
                Console.WriteLine("Invalid Entry\n");

            Console.WriteLine("");

            if (string.Equals(userInput, "print"))
            {
                Console.Write("Enter the ID for the flight you want to print: ");
                string FlightId = Console.ReadLine();
                if (FlightId != null)
                {
                    flighter.getFlightManifest(FlightId);
                }
                else
                {
                    Console.WriteLine("Invalid Entry\n");
                }
            }

            else if (!string.Equals(userInput, "quit"))
            {
                Console.WriteLine("*********************************************************************************************");
                System.Environment.Exit(1);
            }
            else
                Console.WriteLine("Invalid Entry\n");
            Console.WriteLine("*********************************************************************************************");

        } while (!string.Equals(userInput, "quit"));


        Console.WriteLine("*********************************************************************************************");
        System.Environment.Exit(1);
        return;

    }

    public void AccountingManagerCli(SWE_Project.AccountingManager accountant)
    {

        Console.WriteLine("*********************************************************************************************");
        Console.WriteLine("Welcome Back " + accountant.FName + " " + accountant.LName + "!\n");
        string user = accountant.UserId; // Temp
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To select a plane to get the profit of, enter profit.");
            Console.WriteLine("To get the profit of the whole company, enter total.");
            Console.WriteLine("To create an account for a fellow worker, enter account.");
            Console.WriteLine("To exit the marketing manager portal, enter quit.\n");


            userInput = Console.ReadLine();
            if (userInput != null)
                userInput = userInput.ToLower();
            else
                Console.WriteLine("Invalid Entry\n");

            Console.WriteLine("");

            if (string.Equals(userInput, "profit"))
            {
                Console.Write("Enter the ID for the flight you want to get the profit of: ");
                string FlightId = Console.ReadLine();
                if (FlightId != null)
                {
                    accountant.getFlightProfit(FlightId);
                }
                else
                {
                    Console.WriteLine("Invalid Entry\n");
                }
            }
            else if (string.Equals(userInput, "total"))
            {
                accountant.getTotalProfit();
            }
            else if (string.Equals(userInput, "quit"))
            {
                Console.WriteLine("*********************************************************************************************");
                System.Environment.Exit(1);
            }
            else
                Console.WriteLine("Invalid Entry\n");
            Console.WriteLine("*********************************************************************************************");

        } while (!string.Equals(userInput, "quit"));


        Console.WriteLine("*********************************************************************************************");
        System.Environment.Exit(1);

        return;
    }
}


class Program
{
    static void Main(String[] args)
    {
        Globals.databasePath = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + @"\AirportInfo.xlsx"); // store excel file in debug so it can be grabbed 
        CLICaller caller = new CLICaller();
        int Vr = 0;
        string mainInput;
        string user = "";
        string pass = "";
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        Console.WriteLine("*********************************************************************************************");
        Console.WriteLine("Welcome to MidWest Airlines\n");
        do
        {
            Console.WriteLine("If you already have an account and want to access the app, enter Login");
            Console.WriteLine("To make a new account, enter Create ");
            Console.WriteLine("To exit the application, enter Quit\n");
            mainInput = Console.ReadLine();
            if (mainInput != null)
                mainInput = mainInput.ToLower();

            if (string.Equals(mainInput, "login"))//When login is inputted wait for input of the ID and password send to login function
            {
                user = "";
                pass = "";
                Console.Write("Enter user ID: ");
                user = Console.ReadLine();
                Console.Write("Enter password: ");
                pass = Console.ReadLine();
                Vr = Login(user, pass, workbook);


                if (Vr == 0)
                {
                    Console.WriteLine("Username or Password was incorrrect\n");
                    Console.WriteLine("*********************************************************************************************\n");

                }
            }
            else if (mainInput == "create")//When login in inputted ask for Name, Address, Phone, Age, Card Information, Password and send to CreateAccount function
            {
                int part = 0;
                string fname = "";string lname = ""; string address = ""; string phone = ""; string age = "";string card = ""; string passs = ""; string confir = "";
                do
                {
                    if (part == 0)
                    {
                        Console.Write("Enter First Name: ");
                        fname = Console.ReadLine();
                        part++;
                    }
                    else if (part == 1) {
                        Console.Write("Enter Last Name: ");
                        lname = Console.ReadLine();
                        part++;
                    }
                    else if (part == 2)
                    {
                        Console.Write("Enter Address: ");
                        address = Console.ReadLine();
                        part++;
                    }
                    else if (part == 3)
                    {
                        Console.Write("Enter Phone: ");
                        phone = Console.ReadLine();
                        try
                        {
                            Int32.Parse(phone);
                            part++;
                        }
                        catch (ArgumentNullException)
                        {
                            Console.Write("Please Enter a Value");
                        }
                        catch
                        {
                            Console.Write("Invalid Phone Number");
                        }
                    }
                    else if (part == 4)
                    {
                        Console.Write("Enter Age: ");
                        age = Console.ReadLine();
                        try
                        {
                            Int32.Parse(age);
                            part++;
                        }
                        catch (ArgumentNullException)
                        {
                            Console.Write("Please Enter a Value");
                        }
                        catch 
                        {
                            Console.Write("Invalid age");
                        }
                    }
                    else if (part == 5)
                    {
                        Console.Write("Enter Card Information: ");
                        card = Console.ReadLine();
                        if (card != null)
                        {
                            if (card.Length >= 16)
                            {
                                part++;
                            }
                            else
                            {
                                Console.Write("Invalid Card Number\n");
                            }
                        }
                        else
                        {
                            Console.Write("Invalid Entry\n");
                        }
                    }
                    else if (part == 6)
                    {
                        Console.Write("Enter Password: ");
                        passs = Console.ReadLine();
                        if (passs != null)
                        {
                            part++;
                        }
                        else
                        {
                            Console.WriteLine("Invalid Password");
                        }
                    }
                    else if (part == 7)
                    {
                        Console.Write("Confirm Submission (Y/N)");
                        confir = Console.ReadLine();
                        if(confir != null)
                        {
                            if (confir.ToLower() == "y")
                            {
                                CreateAccount(fname, lname, address, phone, age, card, passs);
                                part++;
                            }
                            else if (confir.ToLower() == "n")
                            {
                                part = 8;
                            }
                            else
                            {
                                Console.WriteLine("");
                                Console.WriteLine("Enter Valid Confirmation (Y/N)\n");
                            }
                        }
                        else
                        {
                            Console.WriteLine("");
                            Console.WriteLine("Enter Valid Confirmation (Y/N)\n");
                        }
                       
                    }

                } while (part != 8);
               
            }
            else if (mainInput == "quit")
            {
                System.Environment.Exit(1);
            }
            else
            {
                Console.WriteLine("Invalid Entry\n");
                Console.WriteLine("*********************************************************************************************\n");

            }

        } while (Vr == 0);
        System.DateTime dateTime = System.DateTime.Now;
        CLICaller cLi = new CLICaller();
        if (user.Length == 6)
        {
            var worksheet = workbook.Worksheet("custList");
            var table = worksheet.Tables.Table(0);
            var idCol = table.Column(1);
            Customer currentUser = new Customer(idCol.Cell(Vr).Value.ToString(),
                idCol.Cell(Vr).CellRight(1).Value.ToString(),
                idCol.Cell(Vr).CellRight(8).GetValue<int>(),
                idCol.Cell(Vr).CellRight(9).Value.ToString(),
                idCol.Cell(Vr).CellRight(10).Value.ToString(),
                idCol.Cell(Vr).CellRight(4).Value.ToString(),
                idCol.Cell(Vr).CellRight(6).GetValue<int>(),
                idCol.Cell(Vr).CellRight(5).Value.ToString());
            cLi.CustomerCli(currentUser);
        }
        else if (user.Length == 5)
        {
            var worksheet = workbook.Worksheet("EmpList");
            var table = worksheet.Tables.Table(0);
            var idCol = table.Column(1);
            string dep = idCol.Cell(Vr).CellRight(2).Value.ToString();
            dep = dep.ToLower();
            if (dep == "marketing")
            {
                MarketingManager currentUser = new MarketingManager(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString());
                workbook.Dispose();
                cLi.marketingManagerCli(currentUser);
            }
            else if (dep == "engineer")
            {
                LoadEngineer currentUser = new LoadEngineer(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString());
                workbook.Dispose();
                cLi.LoadEngineerCli(currentUser);
            }
            else if (dep == "flight")
            {
                FlightManager currentUser = new FlightManager(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString());
                workbook.Dispose();
                cLi.FlightManagerCli(currentUser);
            }
            else if (dep == "accounting")
            {
                AccountingManager currentUser = new AccountingManager(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString());
                workbook.Dispose();
                cLi.AccountingManagerCli(currentUser);
            }
        }
       
    }
    static int Login(string user, string pass, XLWorkbook workbook)
    {
        if (user == null || pass == null)
        {
            Console.WriteLine("Invalid Entry\n");
            return 0;
        }
        int usersRow = 0;
        //var workbook = new XLWorkbook(Globals.databasePath); // Open database
        var worksheet = workbook.Worksheet("custList");
        if (user.Length == 6)
        {
            worksheet = workbook.Worksheet("custList"); // Get list of employees
        }
        else if (user.Length == 5)
        {
            worksheet = workbook.Worksheet("EmpList"); // Get list of Employees
        }
        else
        {
            Console.WriteLine("Failed to login - Invalid credentials\n");
            return 0; //If length is not 8 for customers or 7 for employees than username is invalid so return Q
        }

        var table = worksheet.Tables.Table(0); // Get customer Table
        var totalRows = worksheet.LastRowUsed().RowNumber();
        for (int i = 1; i <= totalRows; i++)
        {
            var usCell = table.Row(i).Cell(1).GetString();//Get row user id
            if (string.Equals(usCell, user))
            {
                byte[] tmpNewHash;
                string SavedPass;
                string checkPass;
                SHA512 shaM = new SHA512Managed();
                var tmpSource = ASCIIEncoding.ASCII.GetBytes(pass);//Turns inputted password into bytes
                tmpNewHash = shaM.ComputeHash(tmpSource);//Hashes the bytes
                checkPass = Encoding.UTF8.GetString(tmpNewHash);//turns it back into a string
                SavedPass = table.Row(i).Cell(2).Value.ToString();
                if (checkPass == SavedPass)//Compares inputed hashed string to hashed string stored in database
                {
                    usersRow = i;//Stores row of users information
                }
                else
                {
                    usersRow = 0;
                }
                break;
            }
        }

        if (usersRow == 0)
            Console.WriteLine("Failed to login - Invalid credentials\n");



        return usersRow;
    }
    static bool CreateAccount(string fname, string lname, string address, string phone, string age, string cardin, string pass)
    {
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        var worksheet = workbook.Worksheet("custList");
        var table = worksheet.Tables.Table(0); // Get customer Table
        var lastRowPos = worksheet.LastRowUsed().RowNumber();
        worksheet.Row(lastRowPos).InsertRowsBelow(1);
        Random rnd = new Random();
        int ranCheck = rnd.Next(0, 900000);
        ranCheck = 999999 - ranCheck;
        int cmp;
        for (int x = 2; x <= lastRowPos; x++)
        {
            cmp = worksheet.Row(x).Cell(1).GetValue<int>();
            if (ranCheck == cmp)
            {
                ranCheck = rnd.Next(0, 900000);
                ranCheck = 999999 - ranCheck;
                x = 1;
            }
        }
        lastRowPos++;
        worksheet.Row(lastRowPos).Cell(1).Value = ranCheck;
        worksheet.Row(lastRowPos).Cell(3).Value = fname;
        worksheet.Row(lastRowPos).Cell(4).Value = lname;
        worksheet.Row(lastRowPos).Cell(5).Value = address;
        worksheet.Row(lastRowPos).Cell(6).Value = phone;
        worksheet.Row(lastRowPos).Cell(7).Value = age;
        worksheet.Row(lastRowPos).Cell(8).Value = 0;
        worksheet.Row(lastRowPos).Cell(9).Value = 0;
        worksheet.Row(lastRowPos).Cell(10).Value = cardin;
        byte[] tmpSource;
        byte[] tmpHash;
        String byteholder;
        SHA512 shaM = new SHA512Managed();
        tmpSource = ASCIIEncoding.ASCII.GetBytes(pass);
        tmpHash = shaM.ComputeHash(tmpSource);
        byteholder = Encoding.UTF8.GetString(tmpHash);
        worksheet.Row(lastRowPos).Cell(2).Value = byteholder;
        workbook.SaveAs(Globals.databasePath);
        workbook.Dispose();
        Console.WriteLine($"Your User ID is: '{ranCheck}'");
        return true;
    }

}