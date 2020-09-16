using System;

namespace ReportWriter
{
    class Program
    {
        static void Main(string[] args)
        {
            var lineRow1 = new LineRow();
            var lineRow2 = new LineRow();
            lineRow1.LineText = "My New Doc";
            lineRow2.LineText = "Another Line";
            var section = new Section();
            section.SectionTitle = "Turnips";
            section.LineRows.Add(lineRow1);
            section.LineRows.Add(lineRow2);
            var collater = new Collater(@"D:\Harvey\OneDrive\IT\MyDoc.docx");
            collater.WriteSection(section);
        }
    }
}
