using actor_interface;
using System.Runtime.CompilerServices;
using ClosedXML;
using ClosedXML.Excel;
using System.Collections;

// Class for global variables following c# standards
public class Globals
{
    public static string databasePath = "";

}
// The Location class stores the state and airport
public class Location
{
    public string state {  get; }
    public string airport {  get; }
    public Location(string state, string airport)
    {
        this.state = state;
        this.airport = airport;

    }
}

// The Flight class holds all the information about a flight as well as methods to calculate distance, cost, and points
public class Flight
{
    Location FlightFrom;
    Location FlightTo;
    DateTime FlightTime;
    int FlightId;
    int PlaneType { get; set; }
    float Distance;
    int PointsGenerated;
    Decimal Price { get; set; } // Using Decimal class made to deal with money cause floats and doubles loose precision over calculations
    //+passengers: List<Customer>

    public Flight(int FlightId,Location FlightFrom, Location FlightTo, DateTime FlightTime)
    {
        this.FlightId = FlightId;
        this.FlightFrom = FlightFrom;
        this.FlightTo = FlightTo;
        this.FlightTime = FlightTime;

        CalculateDistance();
        CalculatePrice();
        CalculatePoints();
    }

    void CalculateDistance()
    {
        int CalculatedDistance = 0;

        // Grab from excel based on location

        this.Distance = CalculatedDistance;
    }

    void CalculatePrice()
    {
        Decimal FixedCost = 50;
        Decimal DistanceCost = (Decimal)this.Distance * (Decimal)0.12;
        int NumOfSegments = 3; // Need a more explicit measure based on plane type
        Decimal TsaSegmentCost = 8 * NumOfSegments;

        Decimal TotalCost = FixedCost + DistanceCost + TsaSegmentCost; // will need to check if rounding is needed

        this.Price = TotalCost;
    }
    void CalculatePoints()
    {
        Decimal PointsDec = this.Price * 100;
        int Points = (int)PointsDec; // will need to check behavoir

        this.PointsGenerated = Points;
    }

    void GetPath() { } // Need to chat with group about how to handle connecting flights
}

// The load engineer is responsible for creating, editing, and deleting flights3
public class LoadEngineer : Actor
{

    string UserId { get; }
    string Password { get; set; }

    public LoadEngineer(string UserId, string Password)
    {
        this.UserId = UserId;
        this.Password = Password;

    }

    public void CreateFlight(int FlightId,Location DepartingFrom, Location ArrivingAt, DateTime DateTimeInformation)
    {
        Flight newFlight = new(FlightId,DepartingFrom, ArrivingAt, DateTimeInformation);
        try
        {
            var workbook = new XLWorkbook(Globals.databasePath); // Open database
            var worksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet
     
            var table = worksheet.Tables.Table(0); // Get Flight Table

            var listOfData = new ArrayList(); // Making list to feed data into Append data function (IEnumerable)
            listOfData.Add(FlightId);
            listOfData.Add(DepartingFrom.airport);
            listOfData.Add(ArrivingAt.airport);
            listOfData.Add(DateTimeInformation.ToUniversalTime().ToShortDateString());
          
            table.InsertRowsBelow(1); // Put new flight data into list

            for(int i = 0; i < table.LastRow().CellCount(); i++) // Iterrate through last row of table hitting each cell
            {
                table.LastRow().Cell(i + 1).Value = listOfData[i].ToString(); // Change value of cell to list data

            }
            workbook.Save(); // Save changes
            
        }
        catch (FileNotFoundException ex)
        {
            Console.WriteLine(ex.Message);
            return;
        }



    }

    public void EditFlight(int FlightId)
    {
        // Find excel in excel file

        // bring up data




    }

    public void DeleteFlight(int FlightId) { }

    public void CreateAccount(string UserId, string Password)
    {
        throw new NotImplementedException();
    }

    public void Login(string UserId, string Password)
    {
        throw new NotImplementedException();
    }






}


class Program
{
    static void Main(String[] args)
    {
        Globals.databasePath = System.IO.Path.GetFullPath(Directory.GetCurrentDirectory() + @"\AirportInfo.xlsx"); // store excel file in debug so it can be grabbed 
        Console.WriteLine("Hello World");

        LoadEngineer alex = new("12345", "password");

        DateTime dateTime = DateTime.Now;
        Location from = new("texas", "houston airport");
        Location to = new("Nebraska", "Nebraska airport");
        alex.CreateFlight(123,from, to, dateTime);
    }


}




