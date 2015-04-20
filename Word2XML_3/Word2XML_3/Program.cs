using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Serialization;
using NetOffice.WordApi;

namespace Word2XML_3
{
    [XmlType("Field")]
    public class Field
    {
        [XmlAttribute("FieldID")]
        public string FieldID;
        [XmlAttribute("FieldValue")]
        public string FieldValue;
        public Field() { }
        public Field(string fieldID, string fieldValue)
        {
            this.FieldID = fieldID;
            this.FieldValue = fieldValue;
        }
    }
    public static class FieldSerializer
    {
        public static void SerializeObject(this List<Field> fields, string file)
        {
            var serializer = new XmlSerializer(typeof(List<Field>));
            using (var stream = File.OpenWrite(file))
            {
                serializer.Serialize(stream, fields);
            }
        }
    }
    class Program
    {
        private static void outputXMLFile(List<Field> list, string file)
        {
            FieldSerializer.SerializeObject(list, file);
        }
        private static void extractTextFile(List<Field> list, string file)
        {
            using (FileStream s = File.OpenRead(file))
            using (TextReader reader = new StreamReader(s))
            {
                while (reader.Peek() > -1)
                {
                    string line = reader.ReadLine();
                    string[] fieldArray = line.Split('=');
                    cleanUpWhitespaceAndDelimiters(fieldArray);
                    Field field = new Field(fieldArray[0], fieldArray[1]);
                    list.Add(field);
                }
            }
        }
        private static void extractWordFile(List<Field> list, string file)
        {
            string text = wordDocument2String(file);
            string[] lines = text.Split('\r');
            for (int i = 0; i < lines.Length - 1; i++)
            {
                char[] delimiters = { ':', '?','\v' };
                string[] fieldArray = lines[i].Split(delimiters, StringSplitOptions.RemoveEmptyEntries);
                cleanUpWhitespaceAndDelimiters(fieldArray);
                switch (fieldArray[0])
                {
                    case "Best Pizza Toppings":
                        Field f1 = new Field("Best Pizza Toppings", fieldArray[2]);
                        list.Add(f1);
                        break;
                    case "What is your dream job and why":
                        Field f2 = new Field("Dream Job", fieldArray[2]);
                        list.Add(f2);
                        break;
                    case "Favorite Primary Color and Why":
                        Field f3 = new Field("Favorite Primary Color",fieldArray[1]);
                        break;
                    case "What type of vehicle do you drive":
                        Field f4 = new Field("Vehicle", fieldArray[1]);
                        break;
                    default:
                        Field f5 = new Field(fieldArray[0], fieldArray[1]);
                        list.Add(f5);
                        break;
                }
            }

        }
        private static string wordDocument2String(string file)
        {
            NetOffice.WordApi.Application wordApplication = new NetOffice.WordApi.Application();
            NetOffice.WordApi.Document newDocument = wordApplication.Documents.Open(file);
            string txt = newDocument.Content.Text;
            wordApplication.Quit();
            wordApplication.Dispose();
            return txt;
        }
        private static void cleanUpWhitespaceAndDelimiters(string[] s)
        {
            for (int i = 0; i < s.Length; i++)
            {
                s[i] = (s[i].Replace(";", "")).Trim();
            }
        }
        static void Main(string[] args)
        {
            List<Field> _fields = new List<Field>();
            string localPath = Directory.GetCurrentDirectory();
            string wordFile = localPath + @"\" + args[0];
            string xmlFile = localPath + @"\" + args[1];
            //extractTextFile(_fields, wordFile);
            extractWordFile(_fields, wordFile);
            outputXMLFile(_fields, xmlFile);
        }
    }
}
