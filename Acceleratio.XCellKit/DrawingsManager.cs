using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using A14 = DocumentFormat.OpenXml.Office2010.Drawing;
using Drawing = DocumentFormat.OpenXml.Spreadsheet.Drawing;

namespace Acceleratio.XCellKit
{
    public class DrawingsManager
    {
        // open xml koristi mjernu jedinicu naziva EMU
        // prema definiciji za EMU( English Metric Units) postoji 914400 po inchu
        public const int INT_EMUsPerInch = 914400;
        private List<Image> _images = new List<Image>();
        private readonly List<ImageDetails> _imageDetails = new List<ImageDetails>();

        /// <summary>
        /// Sets the image list source that will be used for cell images. set an image in your cell by using the cell.ImageIndex property.
        /// <para />IMPORTANT: Do not use cell images if you expect a high number of rows in the document
        /// this will kill the excel performance if each row has an image when dealing with > 100 000 rows
        /// </summary>        
        public void SetImages(List<Image> images)
        {
            _images = images;
        }

        public void SetImageForCell(ImageDetails details)
        {
            if (_images.Count == 0)
            {
                return;
            }
            _imageDetails.Add(details);
        }

        public void WriteDrawings(WorksheetPart worksheetPart, OpenXmlWriter writer)
        {
            if (!_images.Any() || !_imageDetails.Any())
            {
                return;
            }
            var drawingPartId = "drawingPart1";
            DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>(drawingPartId);
            int imgCounter = 0;
            foreach (var image in _images)
            {
                var imagePart1 = drawingsPart.AddNewPart<ImagePart>("image/png", "image" + imgCounter);

                imgCounter++;
                using (var ms = new MemoryStream())
                {
                    image.Save(ms, ImageFormat.Png);
                    ms.Position = 0;
                    imagePart1.FeedData(ms);
                }
            }
            using (var drawingsWriter = OpenXmlWriter.Create(drawingsPart))
            {
                drawingsWriter.WriteStartElement(new Xdr.WorksheetDrawing(), new List<OpenXmlAttribute>(), new List<KeyValuePair<string, string>>()
                {
                    new KeyValuePair<string, string>("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
                    new KeyValuePair<string, string>("a", "http://schemas.openxmlformats.org/drawingml/2006/main")
                });
                foreach (var imageDetails in _imageDetails)
                {
                    var twoCellAnchor1 = createImageAnchor(imageDetails);
                    drawingsWriter.WriteElement(twoCellAnchor1);
                }
                drawingsWriter.WriteEndElement();
            }

            var idAtt = new OpenXmlAttribute("id", "http://schemas.openxmlformats.org/officeDocument/2006/relationships", drawingPartId);
            var pageMargins1 = new PageMargins() { Left = 0.7D, Right = 0.7D, Top = 0.75D, Bottom = 0.75D, Header = 0.3D, Footer = 0.3D };
            //  PageSetup pageSetup1 = new PageSetup() { PaperSize = (UInt32Value)9U, Orientation = OrientationValues.Portrait};
            writer.WriteElement(pageMargins1);
            //writer.WriteElement(pageSetup1);
            writer.WriteStartElement(new Drawing(), new List<OpenXmlAttribute>() { idAtt });
            writer.WriteEndElement();
        }



        /// <summary>
        /// kovertiranje u English Metric Units (EMU)        
        /// </summary>
        /// <param name="value"></param>
        /// <param name="dpi"></param>
        /// <returns></returns>
        private long convertPixelsToEMUs(int value, float dpi)
        {
            // da bi pixele pretvorili u inche moramo iskoristiti ovu formulu koja je vezana za dpi slike
            return (long)(INT_EMUsPerInch / dpi) * value;
        }

        private uint _pictureCounter;
        private Xdr.OneCellAnchor createImageAnchor(ImageDetails details)
        {
            var image = _images[details.ImageIndex];

            long imgWidth = convertPixelsToEMUs(image.Width, image.HorizontalResolution);
            long imgHeight = convertPixelsToEMUs(image.Height, image.VerticalResolution);

            // ReSharper disable once CompareOfFloatsByEqualityOperator
            if (details.ImageScaleFactor != 0)
            {
                imgWidth = (long)(imgWidth * details.ImageScaleFactor);
                imgHeight = (long)(imgHeight * details.ImageScaleFactor);
            }
            //ovdje zapravo jos ima posla jer su slike mutne i nema indenta
            //no pokazalo se da za 80k redaka slike jednostavno nisu dobro rjesenje u excelu. Excel se brzo izgenerira, ali sto ti to znaci kad ga jedva u excelu gledas
            //tako da ovaj kod jos treba srediti ako cemo koristiti
            Xdr.OneCellAnchor oneCellAnchor = new Xdr.OneCellAnchor();

            Xdr.FromMarker fromMarker1 = new Xdr.FromMarker();
            Xdr.ColumnId columnId1 = new Xdr.ColumnId();
            columnId1.Text = (details.Column - 1).ToString();
            Xdr.ColumnOffset columnOffset1 = new Xdr.ColumnOffset();
            columnOffset1.Text = "0";
            Xdr.RowId rowId1 = new Xdr.RowId();
            rowId1.Text = (details.Row - 1).ToString();
            Xdr.RowOffset rowOffset1 = new Xdr.RowOffset();
            rowOffset1.Text = "0";

            fromMarker1.Append(columnId1);
            fromMarker1.Append(columnOffset1);
            fromMarker1.Append(rowId1);
            fromMarker1.Append(rowOffset1);

            var anchorExtent = new Xdr.Extent() { Cx = imgWidth, Cy = imgWidth };
            //Xdr.ToMarker toMarker1 = new Xdr.ToMarker();
            //Xdr.ColumnId columnId2 = new Xdr.ColumnId();
            //columnId2.Text = (details.Column-1).ToString();
            //Xdr.ColumnOffset columnOffset2 = new Xdr.ColumnOffset();
            //columnOffset2.Text = "203175";
            //Xdr.RowId rowId2 = new Xdr.RowId();
            //rowId2.Text = (details.Row -1).ToString(); ;
            //Xdr.RowOffset rowOffset2 = new Xdr.RowOffset();
            //rowOffset2.Text = "12675";

            //toMarker1.Append(columnId2);
            //toMarker1.Append(columnOffset2);
            //toMarker1.Append(rowId2);
            //toMarker1.Append(rowOffset2);

            Xdr.Picture picture1 = new Xdr.Picture();

            Xdr.NonVisualPictureProperties nonVisualPictureProperties1 = new Xdr.NonVisualPictureProperties();
            UInt32Value id = (UInt32Value)_pictureCounter;
            _pictureCounter++;
            Xdr.NonVisualDrawingProperties nonVisualDrawingProperties1 = new Xdr.NonVisualDrawingProperties() { Id = id, Name = "Picture " + id };

            Xdr.NonVisualPictureDrawingProperties nonVisualPictureDrawingProperties1 = new Xdr.NonVisualPictureDrawingProperties();
            A.PictureLocks pictureLocks1 = new A.PictureLocks() { NoChangeAspect = true };

            nonVisualPictureDrawingProperties1.Append(pictureLocks1);

            nonVisualPictureProperties1.Append(nonVisualDrawingProperties1);
            nonVisualPictureProperties1.Append(nonVisualPictureDrawingProperties1);

            Xdr.BlipFill blipFill1 = new Xdr.BlipFill();

            A.Blip blip1 = new A.Blip() { Embed = "image" + details.ImageIndex };

            blip1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            A.BlipExtensionList blipExtensionList1 = new A.BlipExtensionList();

            A.BlipExtension blipExtension1 = new A.BlipExtension() { Uri = "{28A0092B-C50C-407E-A947-70E740481C1C}" };

            A14.UseLocalDpi useLocalDpi1 = new A14.UseLocalDpi() { Val = false };
            useLocalDpi1.AddNamespaceDeclaration("a14", "http://schemas.microsoft.com/office/drawing/2010/main");

            blipExtension1.Append(useLocalDpi1);

            blipExtensionList1.Append(blipExtension1);

            blip1.Append(blipExtensionList1);

            A.Stretch stretch1 = new A.Stretch();
            A.FillRectangle fillRectangle1 = new A.FillRectangle();

            stretch1.Append(fillRectangle1);

            blipFill1.Append(blip1);
            blipFill1.Append(stretch1);

            Xdr.ShapeProperties shapeProperties1 = new Xdr.ShapeProperties();

            A.Transform2D transform2D1 = new A.Transform2D();
            A.Offset offset1 = new A.Offset() { X = 0, Y = 0 };
            A.Extents extents1 = new A.Extents() { Cx = imgWidth, Cy = imgHeight };

            transform2D1.Append(offset1);
            transform2D1.Append(extents1);

            A.PresetGeometry presetGeometry1 = new A.PresetGeometry() { Preset = A.ShapeTypeValues.Rectangle };
            A.AdjustValueList adjustValueList1 = new A.AdjustValueList();

            presetGeometry1.Append(adjustValueList1);

            shapeProperties1.Append(transform2D1);
            shapeProperties1.Append(presetGeometry1);

            picture1.Append(nonVisualPictureProperties1);
            picture1.Append(blipFill1);
            picture1.Append(shapeProperties1);
            Xdr.ClientData clientData1 = new Xdr.ClientData();

            oneCellAnchor.Append(fromMarker1);
            oneCellAnchor.Append(anchorExtent);
            //twoCellAnchor1.Append(toMarker1);
            oneCellAnchor.Append(picture1);
            oneCellAnchor.Append(clientData1);
            return oneCellAnchor;
        }
    }

    public class ImageDetails
    {
        public int ImageIndex { get; set; }
        public double ImageScaleFactor { get; set; }
        public int Row { get; set; }
        public int Column { get; set; }
        public int Indent { get; set; }
    }
}
