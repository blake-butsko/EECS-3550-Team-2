using SWE_Project;
using System.Runtime.CompilerServices;

public class Flight
{
    Location FlightFrom;
    Location FlightTo;
    DateTime FlightTime;
    int PlaneType { get; set; }
    float Distance;
    int PointsGenerated;
    Decimal Price { get; set; } // Using Decimal class made to deal with money cause floats and doubles loose precision over calculations
    //+passengers: List<Customer>

    Flight(Location FlightFrom, Location FlightTo, DateTime FlightTime)
    {
        this.FlightFrom = FlightFrom;
        this.FlightTo = FlightTo;
        this.FlightTime = FlightTime;
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


public class LoadEngineer : Actor
{

    string UserId { get; }
    string Password { get; set; }

    LoadEngineer(string UserId, string Password)
    {
        this.UserId = UserId;
        this.Password = Password;

    }

    public void CreateFlight(Location DepartingFrom, Location ArrivingAt, DateTime DateTimeInformation) 
    {
        Flight newFlight = new Flight(DepartingFrom, ArrivingAt, DateTimeInformation);
    
    }

    public void EditFlight() { }

    public void DeleteFlight() { }

    public void CreateAccount(string UserId, string Password)
    {
        throw new NotImplementedException();
    }

    public void Login(string UserId, string Password)
    {
        throw new NotImplementedException();
    }






}
public class Location
{
    string state;
    string airport;
    public Location(string state, string airport) 
    { 
        this.state = state;
        this.airport = airport;
        
    }
}


internal class Program
{
    static void Main(String[] args) 
    {
        Console.WriteLine("Hello World");

    }


}




