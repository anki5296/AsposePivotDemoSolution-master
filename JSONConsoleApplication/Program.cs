using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace JSONConsoleApplication
{
    class Program
    {
        static void Main(string[] args)
        {
            //JObject o1 = JObject.Parse(File.ReadAllText(@"C:\Users\Ankita\Source\Repos\AsposePivotDemoSolution-master\JSONConsoleApplication\File1.json"));

            //JArray spendUSD = (JArray)o1["SpendUSD"];

            //Console.WriteLine(spendUSD);

            using (StreamReader r = new StreamReader(@"C:\Users\Ankita\Source\Repos\AsposePivotDemoSolution-master\AsposePivotDemo\File.json"))

            {

                string json = r.ReadToEnd();

                dynamic array = JsonConvert.DeserializeObject(json);

               // array.SpendUSD

                var result = new Dictionary<string, string>();



                foreach (var field in array)

                {

                    result.Add(field.CFDCOL, Convert.ToString(field.FieldValue.value));

                }

                foreach (var item in result)

                {

                    Console.WriteLine(item.Key + "" + item.Value);

                }

            }

            
        }
    }
}
