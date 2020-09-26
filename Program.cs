using System;

namespace ReportWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            var collater = new Collater(@"D:\Harvey\OneDrive\IT\MyDoc.docx");

            var csv = new DataReaderCsv();
            foreach(var t in csv.ReadCSV(@".\Data\FirstData.csv"))
            {
                var section = new Section();
                section.SectionTitle = t.Title;
                var lineRow1 = new LineRow();
                lineRow1.LineText = t.Words;
                var lineRow2 = new LineRow();
                lineRow2.LineText = t.More_Words;
                section.LineRows.Add(lineRow1);
                section.LineRows.Add(lineRow2);
                collater.WriteSection(section);
            }
                            
        }
    }
}
