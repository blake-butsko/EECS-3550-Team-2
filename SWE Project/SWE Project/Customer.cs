using actor_interface;
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
    }
}
