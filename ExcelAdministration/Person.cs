using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MemberAdministration
{
    /// <summary>
    /// The Datamodel for a Person
    /// </summary>
    public class Person
    {
        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String ID { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String Title { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String Surname { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String Name { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String Address { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String Plz { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String State { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String Telephone { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String Mail { get; set; }

        /// <summary>
        /// Getter and Setter (Special for C# - Class Properties) 
        /// </summary>
        public String Amount { get; set; }

        /// <summary>
        /// Creating a Person - Datamodel for easy storage in the Database. 
        /// </summary>
        /// <param name="ID">ID of the Person (Incremented value) </param>
        /// <param name="title">Form of Address</param>
        /// <param name="surname">Surname of the Person</param>
        /// <param name="name">First name of the Person</param>
        /// <param name="address">Actual address of the Person</param>
        /// <param name="plz">ZIP code of the address of the Person</param>
        /// <param name="state">State, Country or Canton of the Person</param>
        /// <param name="telephone">Mobile, Home or similar Phone number</param>
        /// <param name="mail">E-Mail of the Person</param>
        /// <param name="amount">Amount pledged or already payed by the Person</param>
        public Person(String ID, String title, String surname, String name, String address, String plz, String state, String telephone, String mail, String amount)
        {
            this.ID = ID;
            this.Title = title;
            this.Surname = surname;
            this.Name = name;
            this.Address = address;
            this.Plz = plz;
            this.State = state;
            this.Telephone = telephone;
            this.Mail = mail;
            this.Amount = amount;
        }
    }
}
