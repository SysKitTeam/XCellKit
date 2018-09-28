using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Acceleratio.XCellKit
{
    public class SpreadsheetWorkbook
    {
        public SpreadsheetWorkbook() 
        {
            _worksheets = new List<SpreadsheetWorksheet>();
            _tableCount = 1;
        }
        
        private readonly List<SpreadsheetWorksheet> _worksheets; 
        public void AddWorksheet(SpreadsheetWorksheet spreadsheetWorksheet)
        {
            _worksheets.Add(spreadsheetWorksheet);
        }

        private int _tableCount = 0;
        public void Save(string path)
        {
#if DEBUG
            var sw = Stopwatch.StartNew();
#endif
            using (var document = SpreadsheetDocument.Create(path, SpreadsheetDocumentType.Workbook))
            {
                save(document);
            }
#if DEBUG
            sw.Stop();
            Debug.WriteLine(string.Format("Export time:{0}", sw.Elapsed));
#endif
        }

        public void Save(Stream stream)
        {
            using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
            {
                save(document);
            }
        }

        private void save(SpreadsheetDocument document)
        {
             var sheets = new Sheets();
                var workbookPart = document.AddWorkbookPart();
                var workbook = workbookPart.Workbook = new Workbook();
                workbook.Sheets = sheets;
                var stylesManager = new SpreadsheetStylesManager(workbookPart);
                writeWorkSheets(workbookPart, sheets, stylesManager);
        }

        private void writeWorkSheets(WorkbookPart workbookPart, Sheets sheets, SpreadsheetStylesManager stylesManager)
        {
            var sheetCounter = 1;
            foreach (var worksheets in _worksheets)
            {
                var sheetId = "Sheet" + sheetCounter;
                var sheet = new Sheet() { Name = new string(worksheets.Name.Take(30).ToArray()), SheetId = (UInt32Value)(UInt32)sheetCounter, Id = sheetId };
                sheetCounter++;
                sheets.Append(sheet);
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>(sheetId);
                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    worksheets.WriteWorksheet(writer, worksheetPart, stylesManager, ref _tableCount);
                }
            }
        }
    }
}
