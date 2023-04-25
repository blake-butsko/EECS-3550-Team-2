using actor_interface;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    internal class Customer : Actor
    {
        public string UserId { get; }
        private string Password { get; set; }
        public string FName { get; set; }
        public string LName { get; set; }
        public int Points { get; }


        string CreditCardInfo = ""; // Could make into list to hold several cards
        string Email = "";
        string Address = "";
        public int Age = -1;
        string PhoneNumber = "";


      
        public Customer(string userId, string password, int points, string creditCardInfo, string email, string address, int age, string phoneNumber)
        {
            UserId = userId;
            Password = password;
            Points = points;
            CreditCardInfo = creditCardInfo;
            Email = email;
            Address = address;
            Age = age;
            PhoneNumber = phoneNumber;
        }

        public Customer(string UserId, string FName, string LName, int Age)
        {
            this.UserId = UserId;
            this.FName = FName;
            this.LName = LName;
            this.Age = Age;
        }


        public void Login(string UserId, string Password)
        {
            throw new NotImplementedException();
        }

        /*public void ScheduleFlight(string date, string departure, string arrival)
        {
            throw new NotImplementedException();
            // Basic implementation of just checking dates and departure/destination
            // Check inputs against airport list
            // There are none at this date here are the three closest in date that match arrival/departure
            // Have to write code to deal with date and time and how to aprse them - could ask daniel about this
            // 2-44 have no time so start with this 
            var workbook = new XLWorkbook(Globals.databasePath); // Open workbook and worksheet
            var worksheet = workbook.Worksheet("ActiveFlights");

            var table = worksheet.Tables.Table(0);

            var idColumn = table.DataRange.Column(1);
            for (int i = 1; i <= idColumn.CellCount(); i++)
            {





                for (int i = 0; i < worksheet.Tables.Count(); i++) // Go through each table in the sheet
            {
                if (String.Equals(departure, worksheet.Tables.Table(i).Name)) // Get the table that matches the departure location
                {

                    var table = worksheet.Tables.Table(i);

                    for (int j = 1; j <= table.Column(1).CellCount(); j++) // Itterate through all cities in table (Column 1)
                    {

                        if (String.Equals(arrival, table.Column(1).Cell(j).Value.ToString())) // Get the destination from column
                        {

                            return (int)table.Column(1).Cell(j).CellRight(1).Value; // returns distance


                        }
                    }


                }

            }
        }*/
    }
}
