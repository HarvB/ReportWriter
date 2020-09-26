using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;


namespace ReportWriter
{
    public class Collater
    {
        WordprocessingDocument doc;
        public Collater(string fileName)
        {
            doc = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document);
        }
        public void WriteSection(Section section)
        {
            using (doc)
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());
                var para1 = body.AppendChild(new Paragraph());
                var run1 = para1.AppendChild(new Run());
                var runProperties = run1.AppendChild(new RunProperties(new Bold()));
                run1.AppendChild(new Text(section.SectionTitle));
                foreach(var lineRow in section.LineRows)
                {
                    var para = body.AppendChild(new Paragraph());
                    var run = para.AppendChild(new Run());
                    run.AppendChild(new Text(lineRow.LineText));
                }
            }
        }
        public void WriteLineRow()
        {
            
        }
    }
}