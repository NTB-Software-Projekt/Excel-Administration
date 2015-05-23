using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MemberAdministration
{
    public class Person
    {
        public String ID { get; set; }
        public String Title { get; set; }
        public String Surname { get; set; }
        public String Name { get; set; }
        public String Address { get; set; }
        public String Plz { get; set; }
        public String State { get; set; }
        public Int32 Telephone { get; set; }
        public String Mail { get; set; }
        public Int32 Amount { get; set; }

        public Person(String ID, String title, String surname, String name, String address, String plz, String state, String telephone, String mail, String amount)
        {
            this.ID = ID;
            this.Title = title;
            this.Surname = surname;
            this.Name = name;
            this.Address = address;
            this.Plz = plz;
            this.State = state;
            this.Telephone = Int32.Parse(telephone);
            this.Mail = mail;
            this.Amount = Int32.Parse(amount);
        }
    }
}
