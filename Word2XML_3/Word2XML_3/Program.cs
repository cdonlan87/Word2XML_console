using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Serialization;
using Microsoft.Office.Interop;
using NetOffice.WordApi;

namespace Word2XML_3
{
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
        [XmlElement("Responses")]
        public List<CustomField> Responses;
        public CustomEntry() { }
        public CustomEntry(string file)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(file);
            this.Author = getWordDocumentPropertyValue(doc, "Author");
            this.Title = getWordDocumentPropertyValue(doc, "Title");
            this.Trial = getWordDocumentPropertyValue(doc, "Subject");
            this.Responses = ExtractResponses(doc.Content.Text);
            word.Quit();
        }
        public CustomEntry(string author, string title, string trial, List<CustomField> responses)
        {
            this.Author = author;
            this.Title = title;
            this.Trial = trial;
            this.Responses = responses;
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
