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

// Class for global variables following c# standards
public class Globals
{
    public static string databasePath = "";
}


internal class CLICaller
{
    char state;
    public CLICaller() { } // Constructor

    public void CustomerCli() // add customer object here
    {
        Console.WriteLine("*********************************************************************************************");
        string user = "sample name"; // Temp
        Console.WriteLine("\n Welcome Back " + user + "!\n");
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
        string user = "sample name"; // Temp
        Console.WriteLine("\n Welcome Back " + user + "!\n");
        var userInput = "";
        do
        {
            Console.WriteLine("What would you like to do today?");
            Console.WriteLine("To create a flight, enter create.");
            Console.WriteLine("To edit a flight, enter edit.");
            Console.WriteLine("To delete a flight, enter flight.");
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
            }
            else if (string.Equals(userInput, "edit"))
            {

            }
            else if (string.Equals(userInput, "flight"))
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

    public void marketingManagerCli()
    {
        Console.WriteLine("*********************************************************************************************");
        string user = "sample name"; // Temp
        Console.WriteLine("\n Welcome Back " + user + "!\n");
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

    public void FlightManagerCli()
    {
        Console.WriteLine("*********************************************************************************************");
        string user = "sample name"; // Temp
        Console.WriteLine("\n Welcome Back " + user + "!\n");
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

    public void AccountingManagerCli()
    {

        Console.WriteLine("*********************************************************************************************");
        string user = "sample name"; // Temp
        Console.WriteLine("\n Welcome Back " + user + "!\n");
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
        char Vr = 'Q';
        string input;
        string user = "";
        string pass = "";
        do
        {
            Console.WriteLine("Welcome, Enter L to Login, C to Create an Account, or Q to Quit");
            input = Console.ReadLine();
            if (input == "L")
            {
                user = "";
                pass = "";
                Console.Write("Enter username: ");
                user = Console.ReadLine();
                Console.Write("Enter password: ");
                pass = Console.ReadLine();
                Vr = Login(user, pass);
                if (Vr == 'Q')
                {
                    Console.WriteLine("Username or Password was incorrrect");
                }
            }
            else if (input == "C")
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
            else if (input == "Q")
            {
                System.Environment.Exit(1);
            }

        } while (Vr == 'Q');
        System.DateTime dateTime = System.DateTime.Now;
        SWE_Project.Location from = new("Nashville");
        SWE_Project.Location to = new("Cleveland");
        alex.CreateFlight(555, from, to, dateTime);

        SWE_Project.LoadEngineer x = new SWE_Project.LoadEngineer("123","asd");
        x.EditFlight("555");

    }
    static char Login(string user, string pass)
    {
        if (user == null || pass == null)
        {
            return 'Q';
        }
        char Vr = 'Q';
        var workbook = new XLWorkbook(Globals.databasePath); // Open database
        var worksheet = workbook.Worksheet("custList");
        if (user.Length == 6)
        {
            worksheet = workbook.Worksheet("custList"); // Get Flight Manifest sheet
        }
        else if (user.Length == 5)
        {
            worksheet = workbook.Worksheet("EmpList"); // Get Flight Manifest sheet
        }
        else
        {
            return 'Q'; //If length is not 8 for customers or 7 for employees than username is invalid so return Q
        }

        var table = worksheet.Tables.Table(0); // Get customer Table
        var totalRows = worksheet.LastRowUsed().RowNumber();
        for (int i = 1; i <= totalRows; i++)
        {
            var usCell = table.Row(i).Cell(1).GetString();//Get row user id
            if (usCell == user)
            {
                byte[] tmpNewHash;
                byte[] savedHash;
                string SavedPass;
                string checkPass;
                SHA512 shaM = new SHA512Managed();
                var tmpSource = ASCIIEncoding.ASCII.GetBytes(pass);
                tmpNewHash =shaM.ComputeHash(tmpSource);
                checkPass = Encoding.UTF8.GetString(tmpNewHash);
                SavedPass = table.Row(i).Cell(2).Value.ToString();
                //tmpSource = Encoding.UTF8.GetBytes(table.Row(i).Cell(2).Value.ToString());
                //savedHash = shaM.ComputeHash(tmpSource);
                if (checkPass == SavedPass)
                {
                    if (user.Length == 6)
                    {
                        Vr = 'C';
                    }
                    else if (user.Length == 5)
                    {
                        Vr = 'E';
                    }
                }
                else
                {
                    Vr = 'Q';
                }
                break;
            }
        }
        return Vr;
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