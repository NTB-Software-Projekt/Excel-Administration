using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MemberAdministration
{
    class Person
    {

        String ID;
        String title;
        String surname;
        String name;
        String address;
        String plz;
        String state;
        Int32 telephone;
        String mail;
        String amount;

        public String getID()
        {
            return this.ID;
        }

        public String getTitle()
        {
            return this.title;
        }

        public String getSurname()
        {
            return this.surname;
        }

        public String getName()
        {
            return this.name;
        }

        public String getAddress()
        {
            return this.address;
        }

        public String getPlz()
        {
            return this.plz;
        }

        public String getState(){
            return this.state;
        }

        public Int32 getTelephone(){
            return this.telephone;
        }

        public String getEmail(){
            return this.mail;
        }

        public String getAmount()
        {
            return this.amount;
        } 
    }
}
