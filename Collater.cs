using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml;
using System.Collections.Generic;
using System.Text;

namespace ReportWriter
{
    public class Collater
    {
        private WordprocessingDocument _doc;
        public Collater(string fileName)
        {
            _doc = WordprocessingDocument.Create(fileName, WordprocessingDocumentType.Document);
        }
        public void WriteDocument(Sections sections)
        {
                using (_doc)
                {
                    var mainPart = _doc.AddMainDocumentPart();
                    mainPart.Document = new Document();
                    var body = mainPart.Document.AppendChild(new Body());
                    foreach(var section in sections.SectionList)
                    {
                        WriteSection(body, section);
                    }
                    WriteHeader(mainPart);
                    AddSettingsToMainDocumentPart(mainPart);
                }
        }
        private void WriteSection(Body body, Section section)
        {
            var para1 = body.AppendChild(new Paragraph());
            var run1 = para1.AppendChild(new Run());
            var runProperties = run1.AppendChild(new RunProperties(new Bold()));
            run1.AppendChild(new Text(section.SectionTitle));
            var para2 = body.AppendChild(new Paragraph());
            foreach (var lr in section.LineRows)
            {
                WriteLineRow(para2, lr);
            }
        }
        private void WriteLineRow(Paragraph paragraph, LineRow lineRow)
        {
            var run = paragraph.AppendChild(new Run());
            string txt;
            if (lineRow.LineText.EndsWith('.') | lineRow.LineText.EndsWith('?'))
            {
                txt = lineRow.LineText;
            }
            else
            {
                txt = lineRow.LineText.Replace(lineRow.LineText, $"{lineRow.LineText}.");
            }
            run.AppendChild(new Text(txt));
            run.AppendChild(new Break());
        }
        private void WriteHeader(MainDocumentPart mainPart)
        {
            //mainPart.DeleteParts(mainPart.HeaderParts);
            var headerPart = mainPart.AddNewPart<HeaderPart>("rId2");
            var headerPartId = mainPart.GetIdOfPart(headerPart);
            GenerateHeaderPartContent(headerPart);
            IEnumerable<SectionProperties> sections = mainPart.Document.Body.Elements<SectionProperties>();

                foreach (var section in sections)
                {
                    // Delete existing references to headers and footers
                    section.RemoveAllChildren<HeaderReference>();

                    // Create the new header and footer reference node
                    section.PrependChild<HeaderReference>(new HeaderReference() { Id = headerPartId, Type = HeaderFooterValues.Default });
                }
                mainPart.Document.Save();
        }
        private void GenerateHeaderPartContent(HeaderPart part)
        {
            Header header1 = new Header() { MCAttributes = new MarkupCompatibilityAttributes() { Ignorable = "w14 wp14" } };
            header1.AddNamespaceDeclaration("wpc", "http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas");
            header1.AddNamespaceDeclaration("mc", "http://schemas.openxmlformats.org/markup-compatibility/2006");
            header1.AddNamespaceDeclaration("o", "urn:schemas-microsoft-com:office:office");
            header1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
            header1.AddNamespaceDeclaration("m", "http://schemas.openxmlformats.org/officeDocument/2006/math");
            header1.AddNamespaceDeclaration("v", "urn:schemas-microsoft-com:vml");
            header1.AddNamespaceDeclaration("wp14", "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("wp", "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing");
            header1.AddNamespaceDeclaration("w10", "urn:schemas-microsoft-com:office:word");
            header1.AddNamespaceDeclaration("w", "http://schemas.openxmlformats.org/wordprocessingml/2006/main");
            header1.AddNamespaceDeclaration("w14", "http://schemas.microsoft.com/office/word/2010/wordml");
            header1.AddNamespaceDeclaration("wpg", "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup");
            header1.AddNamespaceDeclaration("wpi", "http://schemas.microsoft.com/office/word/2010/wordprocessingInk");
            header1.AddNamespaceDeclaration("wne", "http://schemas.microsoft.com/office/word/2006/wordml");
            header1.AddNamespaceDeclaration("wps", "http://schemas.microsoft.com/office/word/2010/wordprocessingShape");

            Paragraph paragraph1 = new Paragraph() { RsidParagraphAddition = "00164C17", RsidRunAdditionDefault = "00164C17" };

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            ParagraphStyleId paragraphStyleId1 = new ParagraphStyleId() { Val = "Header" };

            paragraphProperties1.Append(paragraphStyleId1);

            Run run1 = new Run();
            Text text1 = new Text();
            text1.Text = "Header";

            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            header1.Append(paragraph1);

            part.Header = header1;
            header1.Save();
        }
        private void AddSettingsToMainDocumentPart(MainDocumentPart part)
        {
            DocumentSettingsPart settingsPart = part.DocumentSettingsPart;
            if (settingsPart==null)
                settingsPart = part.AddNewPart<DocumentSettingsPart>();
                settingsPart.Settings = new Settings(
                new Compatibility( 
                    new CompatibilitySetting() { 
                    Name = new EnumValue<CompatSettingNameValues>
                            (CompatSettingNameValues.CompatibilityMode),
                    Val = new StringValue("15"),
                    Uri = new StringValue
                            ("http://schemas.microsoft.com/office/word")
                }
            )
            );
            settingsPart.Settings.Save();
        }
    }
}