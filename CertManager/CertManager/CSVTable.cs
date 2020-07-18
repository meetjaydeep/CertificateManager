using System;
using System.Collections.Generic;
using System.Linq;
namespace CertManager
{
    public class CSVTable
    {

        public CSVTable(string table, char delimeter)
            : this(table.Split(new string[] { Environment.NewLine }, StringSplitOptions.None), delimeter)
        {
        }

        public CSVTable(string[] lines, char delimeter)
        {
            Columns = lines[0].Split(delimeter);
            Records = lines.ToList().GetRange(1, lines.Length - 1).Select(line => new CSVRecord(line, Columns, delimeter)).ToList();
        }
        public string[] Columns { get; private set; }
        public List<CSVRecord> Records { get; set; }

        public CSVRecord this[int index]
        {
            get { return Records[index]; }
        }
    }
}