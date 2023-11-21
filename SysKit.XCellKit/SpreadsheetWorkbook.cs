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
        delegate void SheetHandler(int ordinal, SpreadsheetWorksheet worksheet, TableIdProvider tableIdProvider);

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

            saveContent(stream);

            stream.Position = startingPosition;

            saveOpenXmlParts(stream);

#if DEBUG
            sw.Stop();
            Debug.WriteLine($"Export time: {sw.Elapsed}");
#endif
        }

        /// <summary>
        /// To avoid a memory bug in openxml: https://github.com/dotnet/Open-XML-SDK/issues/807
        /// we will save the content first without any OpenXml parts.
        /// In this way we can stream data without loading all of the data in memory
        /// </summary>
        private void saveContent(Stream stream)
        {
            using (var package = Package.Open(stream, FileMode.Create, FileAccess.Write))
            {
                using (var document = SpreadsheetDocument.Create(package, SpreadsheetDocumentType.Workbook))
                {
                    saveContent(document);
                }
            }
        }

        /// <summary>
        /// To avoid a memory bug in openxml: https://github.com/dotnet/Open-XML-SDK/issues/807
        /// we will save the content first without any OpenXml parts.
        /// In this way we can stream data without loading all of the data in memory
        /// </summary>
        private void saveContent(SpreadsheetDocument document)
        {
            _sheets = new Sheets();
            var workbookPart = document.AddWorkbookPart();

            _stylesManager = new SpreadsheetStylesManager();
            writeWorkSheets(workbookPart, _sheets, _stylesManager);
        }

        /// <summary>
        /// To avoid a memory bug in openxml: https://github.com/dotnet/Open-XML-SDK/issues/807
        /// we will save openXml parts separately in a separate save.
        /// This will avoid loading large amounts of data in memory
        /// </summary>
        private void saveOpenXmlParts(Stream stream)
        {
            using (var package = Package.Open(stream, FileMode.Open, FileAccess.ReadWrite))
            {
                using (var document = SpreadsheetDocument.Open(package))
                {
                    saveOpenXmlParts(document);
                }
            }
        }

        /// <summary>
        /// To avoid a memory bug in openxml: https://github.com/dotnet/Open-XML-SDK/issues/807
        /// we will save openXml parts separately in a separate save.
        /// This will avoid loading large amounts of data in memory
        /// </summary>
        private void saveOpenXmlParts(SpreadsheetDocument document)
        {
            attachWorkBook(document);
            attachTagsToDocument(document);

            handleWorkSheetsAction((ordinal, spreadsheetWorksheet, tableIdProvider) =>
            {
                var sheetId = "Sheet" + ordinal;
                var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(sheetId);
                spreadsheetWorksheet.AttachAdditionalParts(document.WorkbookPart, worksheetPart, tableIdProvider);
            });
        }

        private void attachTagsToDocument(SpreadsheetDocument document)
        {
            if (_tag != null)
            {
                document.PackageProperties.Keywords = _tag;
            }
        }

        private void attachWorkBook(SpreadsheetDocument document)
        {
            document.WorkbookPart.Workbook = new Workbook();
            document.WorkbookPart.Workbook.Sheets = _sheets;

            _stylesManager.AttachStylesPart(document.WorkbookPart);
        }

        private void writeWorkSheets(WorkbookPart workbookPart, Sheets sheets, SpreadsheetStylesManager stylesManager)
        {
            handleWorkSheetsAction((ordinal, spreadsheetWorksheet, tableIdProvider) =>
            {
                var sheetId = "Sheet" + ordinal;

                var sheet = new Sheet() { Name = new string(spreadsheetWorksheet.Name.Take(30).ToArray()), SheetId = (UInt32Value)(UInt32)ordinal, Id = sheetId };

                sheets.Append(sheet);
                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>(sheetId);
                using (var writer = OpenXmlWriter.Create(worksheetPart))
                {
                    spreadsheetWorksheet.WriteWorksheet(writer, worksheetPart, workbookPart, stylesManager, tableIdProvider);
                }
            });
        }

        private void handleWorkSheetsAction(SheetHandler handler)
        {
            var sheetCounter = 1;
            var tableIdProvider = new TableIdProvider();
            foreach (var worksheet in _worksheets)
            {
                handler(sheetCounter, worksheet, tableIdProvider);
                sheetCounter++;
            }
        }
    }
}
