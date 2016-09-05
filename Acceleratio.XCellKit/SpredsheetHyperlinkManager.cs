using System.Collections.Generic;

namespace Acceleratio.XCellKit
{
    public class SpredsheetHyperlinkManager
    {
        private Dictionary<SpredsheetLocation, SpredsheetHyperLink> _hyperLinks;
        public SpredsheetHyperlinkManager()
        {
            _hyperLinks = new Dictionary<SpredsheetLocation, SpredsheetHyperLink>();
        }

        public void AddHyperlink(SpredsheetLocation location,  SpredsheetHyperLink hyperLink)
        {
            _hyperLinks[location] = hyperLink;
        }

        public Dictionary<SpredsheetLocation, SpredsheetHyperLink> GetHyperlinks()
        {
            return _hyperLinks;
        }  
    }
}
