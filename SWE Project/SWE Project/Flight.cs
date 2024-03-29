﻿using ClosedXML.Excel;
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
        public string FlightFrom;
        public string FlightTo;
        public System.DateTime departTime;
        public System.DateTime arrivalTime;
        public string FlightId { get; }
        public string PlaneType { get; set; }
        float Distance;
        public int capacity { get; set; }
        public int PointsGenerated;
        public List<Customer> passengers = new List<Customer>();
        public Decimal Price { get; set; } // Using Decimal class made to deal with money cause floats and doubles loose precision over calculations
                                          

        public Flight(string FlightId, string FlightFrom, string FlightTo, System.DateTime departTime, System.DateTime arrivalTime)
        {
            this.FlightId = FlightId;
            this.FlightFrom = FlightFrom;
            this.FlightTo = FlightTo;
            this.departTime = departTime;
            this.arrivalTime = arrivalTime;
            // Fill flights with info when created
            CalculateDistance();
            CalculatePrice();
            CalculatePoints();
            GetCapacity();
            PopulateFlight();
           
        }
        public Flight() { }
        // Get capacity of a flight from database
        private void GetCapacity()
        {
            string[] PossiblePlanes = { "737", "757", "767", "787" };
            int[] Capcacities = { 149, 200, 216, 248 };
            var workbook = new XLWorkbook(Globals.databasePath); // Open database
            var flightWorksheet = workbook.Worksheet("ActiveFlights"); // Get Flight Manifest sheet

            var flightTable = flightWorksheet.Tables.Table(0);

            var idColumn = flightTable.DataRange.Column(1);
            bool foundPlaneType = false;
            // Find flight in active flights
            for (int i = 1; i <= idColumn.CellCount(); i++)
            {
                if (String.Equals(idColumn.Cell(i).Value.ToString(), FlightId))
                {
                    this.PlaneType = idColumn.Cell(i).CellRight(5).Value.ToString();
                    if(string.Equals(this.PlaneType, PossiblePlanes[0]))
                    {
                        this.capacity = Capcacities[0];
                    }else if (string.Equals(this.PlaneType, PossiblePlanes[1]))
                    {
                        this.capacity = Capcacities[1];

                    }
                    else if (string.Equals(this.PlaneType, PossiblePlanes[2]))
                    {
                        this.capacity = Capcacities[2];
                    }
                    else if (string.Equals(this.PlaneType, PossiblePlanes[3]))
                    {
                        this.capacity = Capcacities[3];
                    }
                    break;
                }
            }

           

        }

        // Finds distance of a flight from database
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
        // Calculates price of a flight based on passengers, distance, and size of plane
        private void CalculatePrice()
        {
            Decimal FixedCost = 50;
            Decimal DistanceCost = (Decimal)this.Distance * (Decimal)0.12;
            int NumOfSegments = 3; // Need a more explicit measure based on plane type
            Decimal TsaSegmentCost = 8 * NumOfSegments;

            Decimal TotalCost = FixedCost + DistanceCost + TsaSegmentCost; // will need to check if rounding is needed

            // Apply red eye discount
            if (this.departTime.Hour >= 0 || this.departTime.Hour <= 5 || this.arrivalTime.Hour >= 0 || this.arrivalTime.Hour <= 5)
            {
                TotalCost *= (decimal).8;
            }else if((this.departTime.Hour >= 0 && this.departTime.Hour >= 8) || (this.arrivalTime.Hour >= 19 && this.arrivalTime.Hour < 0)) // Apply Off-Peak discount
            {
                TotalCost *= (decimal).9;
            }

            this.Price = TotalCost;
        }
        // Calculate points generated from a flight
        private void CalculatePoints()
        {
            Decimal PointsDec = this.Price * 100;
            int Points = (int)PointsDec; 

            this.PointsGenerated = Points;
        }

        // Fill flight list with passengers
        private void PopulateFlight() // Could add a flag to prevent needless checks
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var custHistWorksheet = workbook.Worksheet("CustHistory");

            var custHistTable = custHistWorksheet.Tables.Table(0);

            var custHistIdColumn = custHistTable.Column(2);
            List<string> userIds = new List<string>();

            // Find passengers on this flight and add their id to list
            for (int i = 1; i <= custHistIdColumn.CellCount(); i++)
            {
                if ((string.Equals(custHistIdColumn.Cell(i).Value.ToString(), FlightId)))
                    userIds.Add(custHistIdColumn.Cell(i).CellLeft(1).Value.ToString());     
            }

            var custWorksheet = workbook.Worksheet("CustList");
            var custTable = custWorksheet.Tables.Table(0);

            var custIdColumn = custTable.Column(1);
            // Find passengers using customer ids and add customer object to list
            for(int i = 1; i <= custIdColumn.CellCount(); i++)
            {
                if(userIds.Contains(custIdColumn.Cell(i).Value.ToString()))
                {
                    Customer cust = new Customer(custIdColumn.Cell(i).Value.ToString(),
                        custIdColumn.Cell(i).CellRight(2).Value.ToString(),
                            custIdColumn.Cell(i).CellRight(3).Value.ToString(),
                                Int32.Parse(custIdColumn.Cell(i).CellRight(6).Value.ToString()));
                    passengers.Add(cust);
                }

            }

        }
    }

}
