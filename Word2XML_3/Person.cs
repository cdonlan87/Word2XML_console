using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;

namespace Word2XML_3
{
    [XmlType("Person")]
    public class Person
    {
        [XmlAttribute("Name")]
        public string Name;
        [XmlElement("Age Bracket")]
        public MyTupleXML<int, int> AgeBracket;
        [XmlAttribute("Date of Birth")]
        public DateTime DateOfBirth;
        [XmlAttribute("Sex")]
        public string Sex;
        [XmlAttribute("ZipCode")]
        public int ZipCode;
        [XmlAttribute("Phone Number")]
        public string TenDigitPhoneNumber;
        [XmlAttribute("Favorite Primary Color")]
        public string FavoritePrimaryColor;
        [XmlElement("Best Pizza Toppings")]
        public List<string> BestPizzaToppings;
        [XmlAttribute("Dream Job")]
        public string DreamJob;
        [XmlAttribute("Vehicle")]
        public string Vehicle;
        public Person(){}
        public Person(string name, MyTupleXML<int,int> ageBracket, DateTime dateOfBirth, string sex, int zipCode,
            string phoneNumber, string favoritePrimaryColor, List<string> bestPizzaToppings, string dreamJob,
            string vehicle)
        {
            this.Name = name;
            this.AgeBracket = ageBracket;
            this.DateOfBirth = dateOfBirth;
            this.Sex = sex;
            this.ZipCode = zipCode;
            this.TenDigitPhoneNumber = phoneNumber;
            this.FavoritePrimaryColor = favoritePrimaryColor;
            this.BestPizzaToppings = bestPizzaToppings;
            this.DreamJob = dreamJob;
            this.Vehicle = vehicle;
        }

    }
}
