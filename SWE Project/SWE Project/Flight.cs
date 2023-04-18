using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
   

    // The Flight class holds all the information about a flight as well as methods to calculate distance, cost, and points
    internal class Flight
    {
        string FlightFrom;
        string FlightTo;
        System.DateTime FlightTime;
        string FlightId;
        public int PlaneType { get; set; }
        float Distance;
        int PointsGenerated;
        public List<string> passengers = new List<string>();
        public Decimal Price { get; set; } // Using Decimal class made to deal with money cause floats and doubles loose precision over calculations
                                    //+passengers: List<Customer>

        public Flight(string FlightId, string FlightFrom, string FlightTo, System.DateTime FlightTime)
        {
            this.FlightId = FlightId;
            this.FlightFrom = FlightFrom;
            this.FlightTo = FlightTo;
            this.FlightTime = FlightTime;

            CalculateDistance();
            CalculatePrice();
            CalculatePoints();
            PopulateFlight();
        }
        public Flight() { }

        private void CalculateDistance()
        {
            int CalculatedDistance = 0;

            var workbook = new XLWorkbook(Globals.databasePath); // Open workbook and worksheet
            var worksheet = workbook.Worksheet("FlightDistance");

            for (int i = 0; i < worksheet.Tables.Count(); i++) // Go through each table in the sheet
            {
                if (String.Equals(FlightFrom, worksheet.Tables.Table(i).Name)) // Get the table that matches the departure location
                {

                    var table = worksheet.Tables.Table(i);

                    for (int j = 1; j <= table.Column(1).CellCount(); j++) // Itterate through all cities in table (Column 1)
                    {

                        if (String.Equals(FlightTo, table.Column(1).Cell(j).Value.ToString())) // Get the destination from column
                        {

                            CalculatedDistance = (int)table.Column(1).Cell(j).CellRight(1).Value; // Grab pre calculated distance
                            this.Distance = CalculatedDistance;

                            return;
                        }
                    }


                }

            }
        }

        private void CalculatePrice()
        {
            Decimal FixedCost = 50;
            Decimal DistanceCost = (Decimal)this.Distance * (Decimal)0.12;
            int NumOfSegments = 3; // Need a more explicit measure based on plane type
            Decimal TsaSegmentCost = 8 * NumOfSegments;

            Decimal TotalCost = FixedCost + DistanceCost + TsaSegmentCost; // will need to check if rounding is needed

            this.Price = TotalCost;
        }
        private void CalculatePoints()
        {
            Decimal PointsDec = this.Price * 100;
            int Points = (int)PointsDec; // will need to check behavoir

            this.PointsGenerated = Points;
        }

        // Fill flight list with passengers
        private void PopulateFlight() // Could add a flag to prevent needless checks
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("CustHistory");

            var table = worksheet.Tables.Table(0);

            var IdColumn = table.Column(2);

            // Find passengers on this flight and add them to flight list
            for(int i = 1; i <= IdColumn.CellCount(); i++)
            {
                if ((string.Equals(IdColumn.Cell(i).Value.ToString(),FlightId)))
                    passengers.Add(IdColumn.Cell(i).CellLeft(1).Value.ToString());
            }

        }
        void GetPath() { } // Need to chat with group about how to handle connecting flights
    }

}
