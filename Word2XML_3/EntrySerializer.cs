using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.IO;

namespace Word2XML_3
{
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
}
