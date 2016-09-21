using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Acceleratio.XCellKit
{
    class SpredsheetCreatePivotTableCache
    {
        public SpredsheetCreatePivotTableCache(PivotTableCacheDefinitionPart part, string relationshipId)
        {
            PivotTableCacheRecordsPart cacheRecord = part.AddNewPart<PivotTableCacheRecordsPart>(relationshipId);
            cacheRecord.PivotCacheRecords = SpredsheetPivotCacheRecords.records;

            part.PivotCacheDefinition = SpreadsheetPivotCacheDefinition.cacheDefinition;
        }
        
        public void ExternalRelationship(PivotTableCacheDefinitionPart part, string relationshipType, Uri externalUri, string id)
        {
            part.AddExternalRelationship(relationshipType, externalUri, id);
        }
    }
}
