using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;

namespace ReportWriter 
{
    public class DataReaderCsv
    {
        public List<Thingy> ReadCSV(string filePath)
        {
            var thingymawhatcha = new List<Thingy>();
            using (var streamreader = new StreamReader(filePath))
            using (var csv = new CsvReader(streamreader))
            {
                thingymawhatcha = csv.GetRecords<Thingy>().ToList();
            }     
            return thingymawhatcha;
        }
    }
}