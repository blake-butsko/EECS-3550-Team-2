using actor_interface;
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
        string user = person.UserId; // Temp
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        var worksheet = workbook.Worksheet("CustList");
        var table = worksheet.Tables.Table(0); // Get customer Table
        var totalRows = worksheet.LastRowUsed().RowNumber();
        for (int i = 1; i <= totalRows; i++)
        {
            var usCell = table.Row(i).Cell(1).GetString();//Get row user id
            if (string.Equals(usCell, user))
            {
                Console.WriteLine("\n Welcome Back " + table.Row(i).Cell(4).GetString() + "!\n");
                break;
            }
        }
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

            }
            else if (string.Equals(userInput, "account"))
            {

            }
            else if(!string.Equals(userInput, "quit"))
                Console.WriteLine("Invalid Entry\n");

        } while (!string.Equals(userInput, "quit"));
        

        Console.WriteLine("*********************************************************************************************");
        return;
    }

    public void LoadEngineerCli(SWE_Project.LoadEngineer engineer) 
    {
        Console.WriteLine("*********************************************************************************************");
        string user = engineer.UserId; // Temp
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        var worksheet = workbook.Worksheet("EmpList");
        var table = worksheet.Tables.Table(0); // Get customer Table
        var totalRows = worksheet.LastRowUsed().RowNumber();
        for (int i = 1; i <= totalRows; i++)
        {
            var usCell = table.Row(i).Cell(1).GetString();//Get row user id
            if (string.Equals(usCell, user))
            {
                Console.WriteLine("\n Welcome Back " + table.Row(i).Cell(4).GetString() + "!\n");
                break;
            }
        }
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To create a flight, enter create.");
            Console.WriteLine("To edit a flight, enter edit.");
            Console.WriteLine("To delete a flight, enter delete.");
            Console.WriteLine("To create an account for a fellow worker, enter account.");
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
                Console.Write("Enter the airport the flight is taking off from: ");//Need to check if this is an actual airport
                string DepartingFrom = Console.ReadLine();
                Console.Write("Enter the airport the flight will be arriving at: ");
                string ArrivingAt = Console.ReadLine();
                Console.Write("Enter the time of departure: ");//Needs to be more complex
                string DepartTime = Console.ReadLine();
                Console.Write("Enter the time of arrival: ");//Needs to be more complex
                string arrivalTime = Console.ReadLine();
                string confIn;
                do
                {
                    Console.Write("Enter Yes or No (Y/N) to confirm submition: ");
                    confIn = Console.ReadLine();
                    if (confIn == "Y")
                    {
                        //engineer.CreateFlight(FlightId, string DepartingFrom, string ArrivingAt, System.DateTime DateTimeInformation)
                    }
                } while (confIn == "y" || confIn == "n");
            }
            else if (string.Equals(userInput, "edit"))
            {
                Console.Write("Enter the ID for the flight you want to edit: ");
                string FlightId = Console.ReadLine();
                if(FlightId != null)
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
            else if (!string.Equals(userInput, "quit"))
                Console.WriteLine("Invalid Entry\n");

        } while (!string.Equals(userInput, "quit"));


        Console.WriteLine("*********************************************************************************************");
        return;

    }

    public void marketingManagerCli(SWE_Project.MarketingManager marketing)
    {
        Console.WriteLine("*********************************************************************************************");
        /*string user = marketing.UserId; // Temp
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        var worksheet = workbook.Worksheet("EmpList");
        var table = worksheet.Tables.Table(0); // Get customer Table
        var totalRows = worksheet.LastRowUsed().RowNumber();
        for (int i = 1; i <= totalRows; i++)
        {
            var usCell = table.Row(i).Cell(1).GetString();//Get row user id
            if (string.Equals(usCell, user))
            {
                Console.WriteLine("\n Welcome Back " + table.Row(i).Cell(4).GetString() + "!\n");
                break;
            }
        }*/
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To select a plane type for a flight, enter plane.");
            Console.WriteLine("To create an account for a fellow worker, enter account.");
            Console.WriteLine("To exit the marketing manager portal, enter quit.\n");


            userInput = Console.ReadLine();
            if (userInput != null)
                userInput = userInput.ToLower();
            else
                Console.WriteLine("Invalid Entry\n");

            Console.WriteLine("");

            if (string.Equals(userInput, "plane"))
            {
               
            }
            else if (string.Equals(userInput, "account"))
            {

            }
            else if (!string.Equals(userInput, "quit"))
                Console.WriteLine("Invalid Entry\n");

        } while (!string.Equals(userInput, "quit"));


        Console.WriteLine("*********************************************************************************************");
        return;


    }

    public void FlightManagerCli(SWE_Project.FlightManager flighter)
    {
        Console.WriteLine("*********************************************************************************************");
        string user = flighter.UserId; // Temp
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        var worksheet = workbook.Worksheet("EmpList");
        var table = worksheet.Tables.Table(0); // Get customer Table
        var totalRows = worksheet.LastRowUsed().RowNumber();
        for (int i = 1; i <= totalRows; i++)
        {
            var usCell = table.Row(i).Cell(1).GetString();//Get row user id
            if (string.Equals(usCell, user))
            {
                Console.WriteLine("\n Welcome Back " + table.Row(i).Cell(4).GetString() + "!\n");
                break;
            }
        }
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To print a flight manifest for a flight, enter print.");
            Console.WriteLine("To create an account for a fellow worker, enter account.");
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
            else if (string.Equals(userInput, "account"))
            {

            }
            else if (!string.Equals(userInput, "quit"))
                Console.WriteLine("Invalid Entry\n");

        } while (!string.Equals(userInput, "quit"));


        Console.WriteLine("*********************************************************************************************");
        return;

    }

    public void AccountingManagerCli(SWE_Project.AccountingManager accountant)
    {

        Console.WriteLine("*********************************************************************************************");
        string user = accountant.UserId; // Temp
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        var worksheet = workbook.Worksheet("EmpList");
        var table = worksheet.Tables.Table(0); // Get customer Table
        var totalRows = worksheet.LastRowUsed().RowNumber();
        for (int i = 1; i <= totalRows; i++)
        {
            var usCell = table.Row(i).Cell(1).GetString();//Get row user id
            if (string.Equals(usCell, user))
            {
                Console.WriteLine("\n Welcome Back " + table.Row(i).Cell(4).GetString() + "!\n");
                break;
            }
        }
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

            }
            else if (!string.Equals(userInput, "quit"))
                Console.WriteLine("Invalid Entry\n");

        } while (!string.Equals(userInput, "quit"));


        Console.WriteLine("*********************************************************************************************");
        return;
    }
}




class Program
{
    static void Main(String[] args)
    {
        Globals.databasePath = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + @"\AirportInfo.xlsx"); // store excel file in debug so it can be grabbed 
        CLICaller caller = new CLICaller();
        Console.WriteLine("Hello World");
        SWE_Project.LoadEngineer alex = new("12345", "password");
        alex.DeleteFlight("1");

        int Vr = 0;
        string mainInput;
        string user = "";
        string pass = "";
        Console.WriteLine("*********************************************************************************************");
        Console.WriteLine("Welcome to Burger King Airlines");
        Console.WriteLine("");
        do
        {
            Console.WriteLine("If you already have an account and want to access the app, enter Login");
            Console.WriteLine("To make a new account, enter Create ");
            Console.WriteLine("To exit the application, enter Quit");
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
                Vr = Login(user, pass);


                if (Vr == 'Q')
                {
                    Console.WriteLine("Username or Password was incorrrect");
                }
            }
            else if (mainInput == "create")//When login in inputted ask for Name, Address, Phone, Age, Card Information, Password and send to CreateAccount function
            {
                Console.Write("Enter First Name: ");
                string fname = Console.ReadLine();
                Console.Write("Enter Last Name: ");
                string lname = Console.ReadLine();
                Console.Write("Enter Address: ");
                string address = Console.ReadLine();
                Console.Write("Enter Phone: ");
                string phone = Console.ReadLine();
                Console.Write("Enter Age: ");
                string age = Console.ReadLine();
                Console.Write("Enter Password: ");
                string passs = Console.ReadLine();
                Console.Write("Confirm Submission (Y/N)");
                if (Console.ReadLine() == "Y" || Console.ReadLine() == "y")
                {
                    CreateAccount(fname, lname, address, phone, age, passs);
                }
            }
            else if (mainInput == "quit")
            {
                System.Environment.Exit(1);
            }
            else
            {
                Console.WriteLine("Invalid Entry\n");
            }

        } while (Vr == 0);
        System.DateTime dateTime = System.DateTime.Now;
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        CLICaller cLi = new CLICaller();
        if (user.Length == 6)
        {
            var worksheet = workbook.Worksheet("custList");
            var table = worksheet.Tables.Table(0);
            var idCol = table.Column(1);
            Customer currentUser = new Customer(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString(), idCol.Cell(Vr).CellRight(2).GetValue<int>(), idCol.Cell(Vr).CellRight(3).Value.ToString(), idCol.Cell(Vr).CellRight(4).Value.ToString(), idCol.Cell(Vr).CellRight(5).Value.ToString(), idCol.Cell(Vr).CellRight(6).GetValue<int>(), idCol.Cell(Vr).CellRight(7).Value.ToString());
            cLi.CustomerCli(currentUser);
        }
        else if(user.Length == 5)
        {
            var worksheet = workbook.Worksheet("EmpList");
            var table = worksheet.Tables.Table(0);
            var idCol = table.Column(1);
            string dep = idCol.Cell(Vr).CellRight(2).Value.ToString();
            dep = dep.ToLower();
            if (dep == "marketing")
            {
                //MarketingManager currentUser = new MarketingManager(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString());
                //cLi.marketingManagerCli(currentUser);
            }
            else if (dep == "engineer")
            {
                LoadEngineer currentUser = new LoadEngineer(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString());
                cLi.LoadEngineerCli(currentUser);
            }
            else if (dep == "flight")
            {
                FlightManager currentUser = new FlightManager(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString());
                cLi.FlightManagerCli(currentUser);
            }
            else if (dep == "accounting")
            {
                AccountingManager currentUser = new AccountingManager(idCol.Cell(Vr).Value.ToString(), idCol.Cell(Vr).CellRight(1).Value.ToString());
                cLi.AccountingManagerCli(currentUser);
            }
        }
        //SWE_Project.Location from = new("Nashville");
        //SWE_Project.Location to = new("Cleveland");
        //alex.CreateFlight(555, from, to, dateTime);

        //SWE_Project.AccountingManager x = new SWE_Project.AccountingManager("123","password");
        SWE_Project.FlightManager Mark = new SWE_Project.FlightManager("123", "password");
        //Mark.getFlightManifest("555");
        //x.getFlightProfit("555");
    }
    static int Login(string user, string pass)
    {
        if (user == null || pass == null)
        {
            return 0;
        }
        int usersRow = 0;
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
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
            return 0; //If length is not 8 for customers or 7 for employees than username is invalid so return Q
        }

        var table = worksheet.Tables.Table(0); // Get customer Table
        var totalRows = worksheet.LastRowUsed().RowNumber();
        for (int i = 1; i <= totalRows; i++)
        {
            var usCell = table.Row(i).Cell(1).GetString();//Get row user id
            if (string.Equals(usCell , user))
            {
                byte[] tmpNewHash;
                byte[] savedHash;
                string SavedPass;
                string checkPass;
                SHA512 shaM = new SHA512Managed();
                var tmpSource = ASCIIEncoding.ASCII.GetBytes(pass);//Turns inputted password into bytes
                tmpNewHash =shaM.ComputeHash(tmpSource);//Hashes the bytes
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
        return usersRow;
    }
    static bool CreateAccount(string fname, string lname, string address, string phone, string age, string pass)
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
            cmp= worksheet.Row(x).Cell(1).GetValue<int>();
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

        byte[] tmpSource;
        byte[] tmpHash;
        String byteholder;
        SHA512 shaM = new SHA512Managed();
        tmpSource = ASCIIEncoding.ASCII.GetBytes(pass);
        tmpHash = shaM.ComputeHash(tmpSource);
        byteholder = Encoding.UTF8.GetString(tmpHash);
        worksheet.Row(lastRowPos).Cell(2).Value = byteholder;
        workbook.SaveAs(Globals.databasePath);
        Console.WriteLine($"Your User ID is: '{ranCheck}'");
        return true;
    }

}