using System.Collections.Generic;

namespace ReportWriter 
{
    public class Section
    {
        public List<LineRow> LineRows;
        public Section()
        {
            LineRows = new List<LineRow>();
        }
        public string SectionTitle { get; set; }
    }
}