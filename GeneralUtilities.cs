using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace Sharetek
{
    public class GeneralUtilities
    {
        

        static GeneralUtilities()
        {
            
        }

      

       

        public static void SerializeToJsonFile(string fileName, Object obj)
        {
            System.IO.File.WriteAllText(fileName, JsonConvert.SerializeObject(obj, Formatting.Indented));
        }

        public static T DeserializeFromJsonFile<T>(string fileName)
        {
            string json_text = System.IO.File.ReadAllText(fileName);
            return JsonConvert.DeserializeObject<T>(json_text);
        }
    }
}
