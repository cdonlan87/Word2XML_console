using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using Microsoft.Office.Interop.Word;

namespace Word2XML_3
{
    [XmlType("Entry")]
    public class CustomEntry
    {
        [XmlAttribute("Author")]
        public string Author;
        [XmlAttribute("Title")]
        public string Title;
        [XmlAttribute("Trial")]
        public string Trial;
        // [XmlElement("Responses")]
        // public List<CustomField> Responses;
        [XmlElement("Name")]
        public string Name;
        [XmlElement("AgeBracket")]
        public MyTupleXML<int, int> AgeBracket;
        [XmlElement("DateOfBirth")]
        public DateTime DateOfBirth;
        [XmlElement("Sex")]
        public string Sex;
        [XmlElement("ZipCode")]
        public int ZipCode;
        [XmlElement("PhoneNumber")]
        public string TenDigitPhoneNumber;
        [XmlElement("FavoritePrimaryColor")]
        public string FavoritePrimaryColor;
        [XmlElement("BestPizzaToppings")]
        public List<string> BestPizzaToppings;
        [XmlElement("DreamJob")]
        public string DreamJob;
        [XmlElement("Vehicle")]
        public string Vehicle;
        public CustomEntry() { }
        public CustomEntry(string file)
        {
            Application word = new Application();
            Document doc = word.Documents.Open(file);
            this.Author = doc.BuiltInDocumentProperties["Author"].Value;
            this.Title = doc.BuiltInDocumentProperties["Title"].Value;
            this.Trial = doc.BuiltInDocumentProperties["Subject"].Value;
            Person p = ExtractResponses(doc);
            word.Quit();
            this.Name = p.Name;
            this.AgeBracket = p.AgeBracket;
            this.DateOfBirth = p.DateOfBirth;
            this.Sex = p.Sex;
            this.ZipCode = p.ZipCode;
            this.TenDigitPhoneNumber = p.TenDigitPhoneNumber;
            this.FavoritePrimaryColor = p.FavoritePrimaryColor;
            this.BestPizzaToppings = p.BestPizzaToppings;
            this.DreamJob = p.DreamJob;
            this.Vehicle = p.Vehicle;
        }
        private static Person ExtractResponses(Microsoft.Office.Interop.Word.Document doc)
        {
            List<ContentControl> ccList = GetAllContentControls(doc);
            // Name
            string name = ccList[0].Range.Text;
            // Age
            string ageRange1 = ccList[1].Range.Text;
            string[] ageRange2 = ageRange1.Split('-');
            int lower = int.Parse(ageRange2[0]);
            int upper = int.Parse(ageRange2[1]);
            MyTupleXML<int, int> ageRange = new Tuple<int, int>(lower, upper);
            // Date of Birth
            DateTime dateOfBirth = DateTime.Parse(ccList[2].Range.Text);
            // Sex
            string sex;
            if (ccList[3].Checked)
                sex = "Male";
            else if (ccList[4].Checked)
                sex = "Female";
            else sex = "No response";
            // ZipCode
            int zipCode = int.Parse(ccList[5].Range.Text);
            // Phone Number
            string phoneNumber = ccList[6].Range.Text;
            // Favorite Primary Color
            string favoritePrimaryColor = ccList[7].Range.Text;
            // Pizza Toppings
            List<string> pizzaToppings = new List<string>();
            if (ccList[8].Checked)
                pizzaToppings.Add("Pepperoni");
            if (ccList[9].Checked)
                pizzaToppings.Add("Cheese");
            if (ccList[10].Checked)
                pizzaToppings.Add("Jalapenos");
            if (ccList[11].Checked)
                pizzaToppings.Add("Mushrooms");
            if (ccList[12].Checked)
                pizzaToppings.Add("Sausage");
            if (ccList[13].Checked)
                pizzaToppings.Add("Chicken");
            if (ccList[14].Checked)
                pizzaToppings.Add("Beef");
            // Dream Job
            string dreamJob = ccList[15].Range.Text;
            // Vehicle
            string vehicle = ccList[16].Range.Text;
            Person p = new Person(name, ageRange, dateOfBirth, sex, zipCode, phoneNumber, favoritePrimaryColor,
                pizzaToppings, dreamJob, vehicle);
            return p;
        }
        private static List<ContentControl> GetAllContentControls(Document doc)
        {
            List<ContentControl> contentControlList = new List<ContentControl>();
            foreach (ContentControl cc in doc.ContentControls)
            {
                contentControlList.Add(cc);
            }
            return contentControlList;
        }
    }
}
