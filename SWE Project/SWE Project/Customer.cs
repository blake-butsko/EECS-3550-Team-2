using actor_interface;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using static ClosedXML.Excel.XLPredefinedFormat;

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

        public void custHistory()
        {
            var workbook = new XLWorkbook(Globals.databasePath);
            var worksheet = workbook.Worksheet("CustHistory");
            var table = worksheet.Tables.Table(0);
            var flightColumn = table.Column(1); //flight id column
            Console.WriteLine("*********************************************************************************************\n");
            Console.WriteLine("Your Flight History\n");
            Console.WriteLine(flightColumn.Cell(1).CellRight(2).Value.ToString()+ " "+ flightColumn.Cell(1).CellRight(3).Value.ToString()+" "+ flightColumn.Cell(1).CellRight(4).Value.ToString()+" "+ flightColumn.Cell(1).CellRight(5).Value.ToString()+" "+ flightColumn.Cell(1).CellRight(6).Value.ToString()+" "+ flightColumn.Cell(1).CellRight(7).Value.ToString()+" "+ flightColumn.Cell(1).CellRight(8).Value.ToString());
            System.DateTime dateTime;
            System.DateTime dateTimeArrive;
            for (int i = 1; i <= flightColumn.CellCount(); i++)
            {
                if (string.Equals((string)flightColumn.Cell(i).Value, UserId))
                {
                    dateTime = System.DateTime.Parse(flightColumn.Cell(i).CellRight(3).Value.ToString());
                    dateTimeArrive = System.DateTime.Parse(flightColumn.Cell(i).CellRight(4).Value.ToString());

                    if (string.Equals((string)flightColumn.Cell(i).CellRight(8).Value, "Card"))
                    {
                        Console.WriteLine(flightColumn.Cell(i).CellRight(1).Value.ToString(), flightColumn.Cell(i).CellRight(2).Value.ToString(), dateTime, dateTimeArrive, flightColumn.Cell(i).CellRight(5).Value.ToString(), flightColumn.Cell(i).CellRight(6).Value.ToString(), " $" , flightColumn.Cell(i).CellRight(7).Value.ToString(), flightColumn.Cell(i).CellRight(8).Value.ToString());

                    }
                    else
                    {
                        Console.WriteLine(flightColumn.Cell(i).CellRight(1).Value.ToString(), flightColumn.Cell(i).CellRight(2).Value.ToString(), dateTime, dateTimeArrive, flightColumn.Cell(i).CellRight(5).Value.ToString(), flightColumn.Cell(i).CellRight(6).Value.ToString(), " $", flightColumn.Cell(i).CellRight(7).Value.ToString(), flightColumn.Cell(i).CellRight(8).Value.ToString());

                    }
                }
            }
        }


        public void Login(string UserId, string Password)
        {
            throw new NotImplementedException();
        }
    }
}
