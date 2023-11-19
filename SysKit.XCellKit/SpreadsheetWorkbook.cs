using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Packaging;
using System.Linq;

namespace SysKit.XCellKit
{
    public class SpreadsheetWorkbook
    {
        private string _tag = null;

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

        public void SetTag(string tag)
        {
            _tag = tag;
        }

        private int _tableCount = 0;
        private Sheets _sheets;
        private SpreadsheetStylesManager _stylesManager;

        public void Save(string path)
        {
            using (var fs = File.Create(path))
            {
                Save(fs);
            }
        }

        public void Save(Stream stream)
        {
            if (!stream.CanSeek)
            {
                throw new NotSupportedException("Non-seekable streams are not supported");
            }

#if DEBUG
            var sw = Stopwatch.StartNew();
#endif
            var startingPosition = stream.Position;

            firstPassSave(stream);

            stream.Position = startingPosition;

            secondPassSave(stream);

#if DEBUG
            sw.Stop();
            Debug.WriteLine($"Export time: {sw.Elapsed}");
#endif
        }

        private void firstPassSave(Stream stream)
        {
            using (var package = Package.Open(stream, FileMode.Create, FileAccess.Write))
            {
                using (var document = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook))
                {
                    firstPassSave(document);
                }
            }
        }


        private void firstPassSave(SpreadsheetDocument document)
        {
            _sheets = new Sheets();
            var workbookPart = document.AddWorkbookPart();

            _stylesManager = new SpreadsheetStylesManager();
            writeWorkSheets(workbookPart, _sheets, _stylesManager);
        }

        private void secondPassSave(Stream stream)
        {
            using (var package = Package.Open(stream, FileMode.Open, FileAccess.ReadWrite))
            {
                using (var document = SpreadsheetDocument.Open(package))
                {
                    secondPassSave(document);
                }
            }
        }

        private void secondPassSave(SpreadsheetDocument document)
        {
            document.WorkbookPart.Workbook = new Workbook();
            document.WorkbookPart.Workbook.Sheets = _sheets;

            _stylesManager.AttachToWorkBook(document.WorkbookPart);

            if (_tag != null)
            {
                document.PackageProperties.Keywords = _tag;
            }

            var sheetCounter = 1;
            _tableCount = 1;
            foreach (var worksheet in _worksheets)
            {
                var sheetId = "Sheet" + sheetCounter;

                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheetId);

                worksheet.AddTableParts(worksheetPart, ref _tableCount);
            }
        }

        private void writeWorkSheets(WorkbookPart workbookPart, Sheets sheets, SpreadsheetStylesManager stylesManager)
        {
            var sheetCounter = 1;
            _tableCount = 1;
            foreach (var worksheet in _worksheets)
            {
                var sheetId = "Sheet" + sheetCounter;
                var sheet = new Sheet() { Name = new string(worksheet.Name.Take(30).ToArray()), SheetId = (UInt32Value)(UInt32)sheetCounter, Id = sheetId };
                sheetCounter++;
                sheets.Append(sheet);
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>(sheetId);
                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    worksheet.WriteWorksheet(writer, worksheetPart, workbookPart, stylesManager, ref _tableCount);
                }
            }
        }
    }
}
