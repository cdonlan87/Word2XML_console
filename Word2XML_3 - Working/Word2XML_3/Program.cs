using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Serialization;
using Microsoft.Office.Interop;

namespace Word2XML_3
{
    public class MyTupleXML<T1, T2>
    {
        MyTupleXML() { }
        public T1 Item1 { get; set; }
        public T2 Item2 { get; set; }
        public static implicit operator MyTupleXML<T1, T2>(Tuple<T1, T2> t)
        {
            return new MyTupleXML<T1, T2>()
            {
                Item1 = t.Item1,
                Item2 = t.Item2
            };
        }
        public static implicit operator Tuple<T1, T2>(MyTupleXML<T1, T2> t)
        {
            return Tuple.Create(t.Item1, t.Item2);
        }
    }
    [XmlType("Field")]
    public class CustomField
    {
        [XmlAttribute("FieldID")]
        public string FieldID;
        [XmlAttribute("FieldValue")]
        public string FieldValue;
        public CustomField() { }
        public CustomField(string fieldID, string fieldValue)
        {
            this.FieldID = fieldID;
            this.FieldValue = fieldValue;
        }
    }
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
        [XmlElement("Age Bracket")]
        public MyTupleXML<int, int> AgeBracket;
        [XmlElement("Date of Birth")]
        public DateTime DateOfBirth;
        [XmlElement("Sex")]
        public string Sex;
        [XmlElement("ZipCode")]
        public int ZipCode;
        [XmlElement("Phone Number")]
        public string TenDigitPhoneNumber;
        [XmlElement("Favorite Primary Color")]
        public string FavoritePrimaryColor;
        [XmlElement("Best Pizza Toppings")]
        public List<string> BestPizzaToppings;
        [XmlElement("Dream Job")]
        public string DreamJob;
        [XmlElement("Vehicle")]
        public string Vehicle;
        public CustomEntry() { }
        public CustomEntry(string file)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(file);
            this.Author = getWordDocumentPropertyValue(doc, "Author");
            this.Title = getWordDocumentPropertyValue(doc, "Title");
            this.Trial = getWordDocumentPropertyValue(doc, "Subject");
            //this.Responses = ExtractResponses(doc.Content.Text);
            Person p = ExtractResponses2(doc);
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
            word.Quit();
        }
        private static List<Microsoft.Office.Interop.Word.ContentControl> GetAllContentControls(Microsoft.Office.Interop.Word.Document doc)
        {
            List<Microsoft.Office.Interop.Word.ContentControl> contentControlList = new List<Microsoft.Office.Interop.Word.ContentControl>();
            Microsoft.Office.Interop.Word.Range rangeStory;
            foreach (Microsoft.Office.Interop.Word.Range range in doc.StoryRanges)
            {
                rangeStory = range;
                do
                {
                    try
                    {
                        foreach (Microsoft.Office.Interop.Word.ContentControl cc in rangeStory.ContentControls)
                        {
                            contentControlList.Add(cc);
                        }
                        foreach (Microsoft.Office.Interop.Word.Shape shapeRange in rangeStory.ShapeRange)
                        {
                            foreach (Microsoft.Office.Interop.Word.ContentControl cc in shapeRange.TextFrame.TextRange.ContentControls)
                            {
                                contentControlList.Add(cc);
                            }
                        }
                    }
                    catch { }
                    rangeStory = rangeStory.NextStoryRange;
                }
                while (rangeStory != null);
            }
            return contentControlList;
        }
        private static Person ExtractResponses2(Microsoft.Office.Interop.Word.Document doc)
        {
            List<Microsoft.Office.Interop.Word.ContentControl> ccList = GetAllContentControls(doc);
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
            Person p = new Person(name, ageRange,dateOfBirth,sex,zipCode,phoneNumber,favoritePrimaryColor, 
                pizzaToppings,dreamJob,vehicle);
            return p;
        }
        private static List<CustomField> ExtractResponses(string text)
        {
            List<CustomField> list = new List<CustomField>();
            string[] lines = text.Split('\r');
            for (int i = 0; i < lines.Length - 1; i++)
            {

                char[] delimiters = { ':', '?', '\v' };
                string[] fieldArray = lines[i].Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                cleanUpWhitespaceAndDelimiters(fieldArray);
                switch (fieldArray[0])
                {
                    case "Best Pizza Toppings":
                        CustomField f1 = new CustomField("Best Pizza Toppings", fieldArray[2]);
                        list.Add(f1);
                        break;
                    case "What is your dream job and why":
                        CustomField f2 = new CustomField("Dream Job", fieldArray[2]);
                        list.Add(f2);
                        break;
                    case "Favorite Primary Color and Why":
                        CustomField f3 = new CustomField("Favorite Primary Color", fieldArray[1]);
                        break;
                    case "What type of vehicle do you drive":
                        CustomField f4 = new CustomField("Vehicle", fieldArray[1]);
                        break;
                    default:
                        CustomField f5 = new CustomField(fieldArray[0], fieldArray[1]);
                        list.Add(f5);
                        break;
                }
            }
            return list;
        }
        private static string getWordDocumentPropertyValue(Microsoft.Office.Interop.Word.Document document, string propertyName)
        {
            object builtInProperties = document.BuiltInDocumentProperties;
            Type builtInPropertiesType = builtInProperties.GetType();
            object property = builtInPropertiesType.InvokeMember("Item", System.Reflection.BindingFlags.GetProperty, null, builtInProperties, new object[] { propertyName });
            Type propertyType = property.GetType();
            object propertyValue = propertyType.InvokeMember("Value", System.Reflection.BindingFlags.GetProperty, null, property, new object[] { });
            return propertyValue.ToString();
        }
        private static void cleanUpWhitespaceAndDelimiters(string[] s)
        {
            for (int i = 0; i < s.Length; i++)
            {
                s[i] = (s[i].Replace(";", "")).Trim();
            }
        }
    }
    public static class EntrySerializer
    {
        public static void SerializeObject(this CustomEntry entry, string file)
        {
            var serializer = new XmlSerializer(typeof(CustomEntry));
            using (var stream = File.OpenWrite(file))
            {
                serializer.Serialize(stream, entry);
            }
        }
    }
    class Program
    {  
        static void Main(string[] args)
        {
            List<CustomField> _fields = new List<CustomField>();
            string localPath = Directory.GetCurrentDirectory();
            string wordFile = localPath + @"\" + args[0];
            string xmlFile = localPath + @"\" + args[1];
            //extractTextFile(_fields, wordFile);
            CustomEntry entry = new CustomEntry(wordFile);
            EntrySerializer.SerializeObject(entry,xmlFile);
        }
    }
}
