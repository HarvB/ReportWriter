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
            using (this.doc)
            {
                var mainPart = doc.AddMainDocumentPart();
                mainPart.Document = new Document();
                var body = mainPart.Document.AppendChild(new Body());
                foreach(var lineRow in section.LineRows)
                {
                var para = body.AppendChild(new Paragraph());
                var run = para.AppendChild(new Run());
                run.AppendChild(new Text(lineRow.LineText));
                }
            }
        }
    }
}