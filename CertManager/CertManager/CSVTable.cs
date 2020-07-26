using System;
using System.Collections.Generic;
using System.Linq;
namespace CertManager
{
    public class CSVTable
    {
        public CSVTable(string table, char delimeter, bool allValuesRequired)
            : this(table.Split(new string[] { Environment.NewLine }, StringSplitOptions.None), delimeter, allValuesRequired)
        {
            
        }

        public CSVTable(string[] lines, char delimeter, bool allValuesRequired)
        {
            Columns = lines[0].Split(delimeter);


            if (allValuesRequired)
            {
                Records = new List<CSVRecord>();

                foreach (var line in lines.Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(line) || delimeter.ToString().Equals(line))
                    {
                        continue;
                    }

                    CSVRecord csvRecord = new CSVRecord(line, Columns, delimeter);
                    if (csvRecord.Count == Columns.Length)
                    {
                        Records.Add(csvRecord);
                    }
                }
            }
            else
            {
                Records = lines.ToList().GetRange(1, lines.Length - 1)
                    .Where(line => !string.IsNullOrWhiteSpace(line) && !delimeter.ToString().Equals(line))
                    .Select(line => new CSVRecord(line, Columns, delimeter)).ToList();
            }
        }

        public string[] Columns { get; private set; }
        public List<CSVRecord> Records { get; set; }

        public CSVRecord this[int index]
        {
            get { return Records[index]; }
        }
    }
}