using actor_interface;
using System.Runtime.CompilerServices;
using ClosedXML;
using ClosedXML.Excel;
using System.Collections;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.VariantTypes;
using static ClosedXML.Excel.XLPredefinedFormat;

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


    }


}




