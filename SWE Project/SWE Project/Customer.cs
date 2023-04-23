using actor_interface;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SWE_Project
{
    internal class Customer : Actor
    {
        public string UserId { get; }
        private string Password { get; set; }
        int Points { get; }

        string CreditCardInfo; // Could make into list to hold several cards
        string Email;
        string Address;
        int Age;
        string PhoneNumber;


      
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

        public Customer(int UserId)
        {

        }


        public void Login(string UserId, string Password)
        {
            throw new NotImplementedException();
        }
    }
}
