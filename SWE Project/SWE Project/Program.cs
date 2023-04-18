using actor_interface;
using System.Runtime.CompilerServices;
using ClosedXML;
using ClosedXML.Excel;
using System.Collections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.VariantTypes;
using static ClosedXML.Excel.XLPredefinedFormat;
using System.Text;
using System.Security.Cryptography;

// Class for global variables following c# standards
public class Globals
{
    public static string databasePath = "";

}




class Program
{
    static void Main(String[] args)
    {
        Globals.databasePath = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + @"\AirportInfo.xlsx"); // store excel file in debug so it can be grabbed 
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
            var usCell = table.Row(i).Cell(1).GetString();//Get hashed pass
            if (usCell == user)
            {
                var tmpSource = ASCIIEncoding.ASCII.GetBytes(pass);
                byte[] tmpNewHash;
                byte[] savedHash;
                tmpNewHash = new MD5CryptoServiceProvider().ComputeHash(tmpSource);
                tmpSource = ASCIIEncoding.ASCII.GetBytes(table.Row(i).Cell(2).GetString());
                savedHash = new MD5CryptoServiceProvider().ComputeHash(tmpSource);
                bool bEqual = false;
                if (tmpNewHash.Length == savedHash.Length)//Compared stored hash with inputed password
                {
                    int buf = 0;
                    while ((buf < tmpNewHash.Length) && (tmpNewHash[buf] == savedHash[buf]))
                    {
                        buf += 1;
                    }
                    if (buf == tmpNewHash.Length)
                    {
                        bEqual = true;
                    }
                }
                if (bEqual)
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
        int id = worksheet.Row(lastRowPos).Cell(1).GetValue<int>();
        id++;
        lastRowPos++;
        worksheet.Row(lastRowPos).Cell(1).Value = id;
        worksheet.Row(lastRowPos).Cell(3).Value = fname;
        worksheet.Row(lastRowPos).Cell(4).Value = lname;
        worksheet.Row(lastRowPos).Cell(5).Value = address;
        worksheet.Row(lastRowPos).Cell(6).Value = phone;
        worksheet.Row(lastRowPos).Cell(7).Value = age;
        worksheet.Row(lastRowPos).Cell(8).Value = 0;
        worksheet.Row(lastRowPos).Cell(9).Value = 0;

        byte[] tmpSource;
        byte[] tmpHash;
        tmpSource = ASCIIEncoding.ASCII.GetBytes(pass);
        tmpHash = new MD5CryptoServiceProvider().ComputeHash(tmpSource);
        int i;
        StringBuilder sOutput = new StringBuilder(tmpHash.Length);
        for (i = 0; i < tmpHash.Length; i++)
        {
            sOutput.Append(tmpHash[i].ToString("X2"));
        }
        worksheet.Row(lastRowPos).Cell(2).Value = sOutput.ToString();
        workbook.SaveAs(Globals.databasePath);
        return true;
    }

}




