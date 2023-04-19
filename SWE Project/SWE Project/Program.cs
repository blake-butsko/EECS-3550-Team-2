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




class Program
{
    static void Main(String[] args)
    {
        Globals.databasePath = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + @"\AirportInfo.xlsx"); // store excel file in debug so it can be grabbed 
        Console.WriteLine("Hello World, This is the stuff");
        /*SWE_Project.LoadEngineer alex = new("12345", "password");

        System.DateTime dateTime = System.DateTime.Now;
        SWE_Project.Location from = new("Nashville");
        SWE_Project.Location to = new("Cleveland");
        alex.CreateFlight(888, from, to, dateTime);*/

        SWE_Project.MarketingManager benjamin = new("12345", "password");
        benjamin.ChoosePlane(888);

    }


}




