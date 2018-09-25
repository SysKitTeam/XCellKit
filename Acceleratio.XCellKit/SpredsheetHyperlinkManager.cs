using System.Collections.Generic;

namespace Acceleratio.XCellKit
{
    public class SpreadsheetHyperlinkManager
    {
        private Dictionary<SpreadsheetLocation, SpreadsheetHyperLink> _hyperLinks;
        public SpreadsheetHyperlinkManager()
        {
            _hyperLinks = new Dictionary<SpreadsheetLocation, SpreadsheetHyperLink>();
        }

        public void AddHyperlink(SpreadsheetLocation location,  SpreadsheetHyperLink hyperLink)
        {
            _hyperLinks[location] = hyperLink;
        }

        public Dictionary<SpreadsheetLocation, SpreadsheetHyperLink> GetHyperlinks()
        {
            return _hyperLinks;
        }  
    }
}
