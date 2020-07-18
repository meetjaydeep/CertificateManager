using System.Collections.Generic;
namespace CertManager
{
    public class CSVRecord : Dictionary<string, string>
    {
        public CSVRecord(string line, string[] keys, char delimeter)
            : base()
        {
            var lista = line.Split(delimeter);
            for (int i = 0; i < lista.Length; i++)
            {
                Add(keys[i], lista[i]);
            }
        }
    }
}