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
    class Program
    {  
        static void Main(string[] args)
        {
            string localPath = Directory.GetCurrentDirectory();
            string wordFile = localPath + @"\" + args[0];
            string xmlFile = localPath + @"\" + args[1];
            //extractTextFile(_fields, wordFile);
            CustomEntry entry = new CustomEntry(wordFile);
            EntrySerializer.SerializeObject(entry,xmlFile);
        }
    }
}
