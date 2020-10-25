using System;
using System.IO;
using YamlDotNet.RepresentationModel;

namespace ReportWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            var yaml = @".\Config\config.yaml";
            using var sr = new StreamReader(yaml);
            var serial = new YamlStream();
            serial.Load(sr);
            var outfile = serial.Documents[0].RootNode["outfile"];
            var infile = serial.Documents[0].RootNode["infile"];
            var collater = new Collater(outfile.ToString());
            var csv = new DataReaderCsv();
            var sections = new Sections();
            foreach(var t in csv.ReadCSV(infile.ToString()))
            {
                var section = new Section();
                section.SectionTitle = t.Title;
                var lineRow1 = new LineRow();
                lineRow1.LineText = t.Words;
                var lineRow2 = new LineRow();
                lineRow2.LineText = t.More_Words;
                section.LineRows.Add(lineRow1);
                section.LineRows.Add(lineRow2);
                sections.SectionList.Add(section);
            }
            collater.WriteDocument(sections);

        }
    }
}
