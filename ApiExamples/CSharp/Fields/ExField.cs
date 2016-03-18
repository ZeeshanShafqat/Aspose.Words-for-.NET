// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.Globalization;
using System.Text.RegularExpressions;
using System.Threading;

using Aspose.Words;
using Aspose.Words.Fields;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExField : ApiExampleBase
    {
        [Test]
        public void UpdateToc()
        {
            Document doc = new Document();

            //ExStart
            //ExId:UpdateTOC
            //ExSummary:Shows how to completely rebuild TOC fields in the document by invoking field update.
            doc.UpdateFields();
            //ExEnd
        }

        [Test]
        public void GetFieldType()
        {
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            //ExStart
            //ExFor:FieldType
            //ExFor:FieldChar
            //ExFor:FieldChar.FieldType
            //ExSummary:Shows how to find the type of field that is represented by a node which is derived from FieldChar.
            FieldChar fieldStart = (FieldChar)doc.GetChild(NodeType.FieldStart, 0, true);
            FieldType type = fieldStart.FieldType;
            //ExEnd
        }

        [Test]
        public void GetFieldFromDocument()
        {
            //ExStart
            //ExFor:FieldChar.GetField
            //ExId:GetField
            //ExSummary:Demonstrates how to retrieve the field class from an existing FieldStart node in the document.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            FieldStart fieldStart = (FieldStart)doc.GetChild(NodeType.FieldStart, 0, true);

            // Retrieve the facade object which represents the field in the document.
            Field field = fieldStart.GetField();
            
            Console.WriteLine("Field code:" + field.GetFieldCode());
            Console.WriteLine("Field result: " + field.Result);
            Console.WriteLine("Is locked: " + field.IsLocked);

            // This updates only this field in the document.
            field.Update();
            //ExEnd
        }

        [Test]
        public void GetFieldFromFieldCollection()
        {
            //ExStart
            //ExId:GetFieldFromFieldCollection
            //ExSummary:Demonstrates how to retrieve a field using the range of a node.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            Field field = doc.Range.Fields[0];

            // This should be the first field in the document - a TOC field.
            Console.WriteLine(field.Type);
            //ExEnd
        }

        [Test]
        public void InsertTcField()
        {
            //ExStart
            //ExId:InsertTCField
            //ExSummary:Shows how to insert a TC field into the document using DocumentBuilder.
            // Create a blank document.
            Document doc = new Document();

            // Create a document builder to insert content with.
            DocumentBuilder builder = new DocumentBuilder(doc);

            // Insert a TC field at the current document builder position.
            builder.InsertField("TC \"Entry Text\" \\f t");
            //ExEnd
        }

        [Test]
        public void ChangeLocale()
        {
            // Create a blank document.
            Document doc = new Document();
            DocumentBuilder b = new DocumentBuilder(doc);
            b.InsertField("MERGEFIELD Date");

            //ExStart
            //ExId:ChangeCurrentCulture
            //ExSummary:Shows how to change the culture used in formatting fields during update.
            // Store the current culture so it can be set back once mail merge is complete.
            CultureInfo currentCulture = Thread.CurrentThread.CurrentCulture;
            // Set to German language so dates and numbers are formatted using this culture during mail merge.
            Thread.CurrentThread.CurrentCulture = new CultureInfo("de-DE");

            // Execute mail merge.
            doc.MailMerge.Execute(new string[] { "Date" }, new object[] { DateTime.Now });

            // Restore the original culture.
            Thread.CurrentThread.CurrentCulture = currentCulture;
            //ExEnd

            doc.Save(MyDir + "Field.ChangeLocale Out.doc");
        }

        [Test]
        public void RemoveTocFromDocument()
        {
            //ExStart
            //ExFor:CompositeNode.GetChildNodes(NodeType, Boolean)
            //ExId:RemoveTableOfContents
            //ExSummary:Demonstrates how to remove a specified TOC from a document.
            // Open a document which contains a TOC.
            Document doc = new Document(MyDir + "Document.TableOfContents.doc");

            // Remove the first TOC from the document.
            Field tocField = doc.Range.Fields[0];
            tocField.Remove();

            // Save the output.
            doc.Save(MyDir + "Document.TableOfContentsRemoveTOC Out.doc");
            //ExEnd
        }

        [Test]
        //ExStart
        //ExId:TCFieldsRangeReplace
        //ExSummary:Shows how to find and insert a TC field at text in a document. 
        public void InsertTcFieldsAtText()
        {
            Document doc = new Document();

            // Insert a TC field which displays "Chapter 1" just before the text "The Beginning" in the document.
            doc.Range.Replace(new Regex("The Beginning"), new InsertTcFieldHandler("Chapter 1", "\\l 1"), false);
        }

        public class InsertTcFieldHandler : IReplacingCallback
        {
            // Store the text and switches to be used for the TC fields.
            private string mFieldText;
            private string mFieldSwitches;

            /// <summary>
            /// The switches to use for each TC field. Can be an empty string or null.
            /// </summary>
            public InsertTcFieldHandler(string switches) : this(string.Empty, switches)
            {
                this.mFieldSwitches = switches;
            }

            /// <summary>
            /// The display text and switches to use for each TC field. Display name can be an empty string or null.
            /// </summary>
            public InsertTcFieldHandler(string text, string switches)
            {
                this.mFieldText = text;
                this.mFieldSwitches = switches;
            }

            ReplaceAction IReplacingCallback.Replacing(ReplacingArgs args)
            {
                // Create a builder to insert the field.
                DocumentBuilder builder = new DocumentBuilder((Document)args.MatchNode.Document);
                // Move to the first node of the match.
                builder.MoveTo(args.MatchNode);

                // If the user specified text to be used in the field as display text then use that, otherwise use the 
                // match string as the display text.
                string insertText;

                if (!string.IsNullOrEmpty(this.mFieldText))
                    insertText = this.mFieldText;
                else
                    insertText = args.Match.Value;

                // Insert the TC field before this node using the specified string as the display text and user defined switches.
                builder.InsertField(string.Format("TC \"{0}\" {1}", insertText, this.mFieldSwitches));

                // We have done what we want so skip replacement.
                return ReplaceAction.Skip;
            }
        }
        //ExEnd

        //ToDo: Need to more info from dev
        [Test]
        public void InsertBarCodeWord2Pdf()
        {
            Document doc = new Document();
            DocumentBuilder builder = new DocumentBuilder(doc);

            FieldBarcode barcode = new FieldBarcode();
            barcode.IsUSPostalAddress = true;
            barcode.PostalAddress = "60629-5113";
            
            doc.Save(MyDir + "123.docx");

            //// Set custom barcode generator
            //doc.FieldOptions.BarcodeGenerator = new CustomBarcodeGenerator();
            //doc.MailMerge.Execute(new string[] { "ZIP4" }, new object[] { "60629-5113" });

            //doc.Save(MyDir + "InsertBarCodeWord2Pdf.docx", SaveFormat.Docx);

            ////MemoryStream dstStream = new MemoryStream();
            ////doc.Save(dstStream, SaveFormat.Docx);

            ////FieldCollection fields = doc.Range.Fields;
            ////foreach (Field field in fields)
            ////{
            ////    if (field.Type == FieldType.FieldDisplayBarcode)
            ////        Assert.IsTrue(field.Separator.NextSibling == field.End);
            ////}

            //doc.Save(MyDir + "InsertBarCodeWord2Pdf.pdf", SaveFormat.Pdf);

            //// Open document
            //Aspose.Pdf.Document pdfDocument = new Aspose.Pdf.Document(MyDir + "InsertBarCodeWord2Pdf_OUT.pdf");
            //// Get values from all fields
            //foreach (Aspose.Pdf.InteractiveFeatures.Forms.Field formField in pdfDocument.Form)
            //{
            //    Console.WriteLine("Field Name : {0} ", formField.PartialName);
            //    Console.WriteLine("Value : {0} ", formField.Value);
            //}
        }

        ///// <summary>
        ///// Sample of custom barcode generator implementation (with underlying Aspose.BarCode module)
        ///// </summary>
        //public class CustomBarcodeGenerator : IBarcodeGenerator
        //{
        //    /// <summary>
        //    /// Converts barcode type from Word to Aspose.BarCode.
        //    /// </summary>
        //    /// <param name="inputCode"></param>
        //    /// <returns></returns>
        //    private static Symbology ConvertBarcodeType(string inputCode)
        //    {
        //        if (inputCode == null)
        //            return (Symbology)int.MinValue;

        //        string type = inputCode.ToUpper();

        //        switch (type)
        //        {
        //            case "QR":
        //                return Symbology.QR;
        //            case "CODE128":
        //                return Symbology.Code128;
        //            case "CODE39":
        //                return Symbology.Code39Standard;
        //            case "EAN8":
        //                return Symbology.EAN8;
        //            case "EAN13":
        //                return Symbology.EAN13;
        //            case "UPCA":
        //                return Symbology.UPCA;
        //            case "UPCE":
        //                return Symbology.UPCE;
        //            case "ITF14":
        //                return Symbology.ITF14;
        //            case "CASE":
        //                break;
        //        }

        //        return (Symbology)int.MinValue;
        //    }

        //    /// <summary>
        //    /// Converts barcode image height from Word units to Aspose.BarCode units.
        //    /// </summary>
        //    /// <param name="heightInTwipsString"></param>
        //    /// <returns></returns>
        //    private static float ConvertSymbolHeight(string heightInTwipsString)
        //    {
        //        // Input value is in 1/1440 inches (twips)
        //        int heightInTwips = int.MinValue;
        //        int.TryParse(heightInTwipsString, out heightInTwips);

        //        if (heightInTwips == int.MinValue)
        //            throw new Exception("Error! Incorrect height - " + heightInTwipsString + ".");

        //        // Convert to mm
        //        return (float)(heightInTwips * 25.4 / 1440);
        //    }

        //    /// <summary>
        //    /// Converts barcode image color from Word to Aspose.BarCode.
        //    /// </summary>
        //    /// <param name="inputColor"></param>
        //    /// <returns></returns>
        //    private static Color ConvertColor(string inputColor)
        //    {
        //        // Input should be from "0x000000" to "0xFFFFFF"
        //        int color = int.MinValue;
        //        int.TryParse(inputColor.Replace("0x", ""), out color);

        //        if (color == int.MinValue)
        //            throw new Exception("Error! Incorrect color - " + inputColor + ".");

        //        return Color.FromArgb(color >> 16, (color & 0xFF00) >> 8, color & 0xFF);

        //        // Backword conversion -
        //        //return string.Format("0x{0,6:X6}", mControl.ForeColor.ToArgb() & 0xFFFFFF);
        //    }

        //    /// <summary>
        //    /// Converts bar code scaling factor from percents to float.
        //    /// </summary>
        //    /// <param name="scalingFactor"></param>
        //    /// <returns></returns>
        //    private static float ConvertScalingFactor(string scalingFactor)
        //    {
        //        bool isParsed = false;
        //        int percents = int.MinValue;
        //        int.TryParse(scalingFactor, out percents);

        //        if (percents != int.MinValue)
        //        {
        //            if (percents >= 10 && percents <= 10000)
        //                isParsed = true;
        //        }

        //        if (!isParsed)
        //            throw new Exception("Error! Incorrect scaling factor - " + scalingFactor + ".");

        //        return percents / 100.0f;
        //    }

        //    /// <summary>
        //    /// Implementation of the GetBarCodeImage() method for IBarCodeGenerator interface.
        //    /// </summary>
        //    /// <param name="parameters"></param>
        //    /// <returns></returns>
        //    public Image GetBarcodeImage(BarcodeParameters parameters)
        //    {
        //        if (parameters.BarcodeType == null || parameters.BarcodeValue == null)
        //            return null;

        //        BarCodeBuilder builder = new BarCodeBuilder();

        //        builder.SymbologyType = ConvertBarcodeType(parameters.BarcodeType);
        //        if (builder.SymbologyType == (Symbology)int.MinValue)
        //            return null;

        //        builder.CodeText = parameters.BarcodeValue;

        //        if (builder.SymbologyType == Symbology.QR)
        //            builder.Display2DText = parameters.BarcodeValue;

        //        if (parameters.ForegroundColor != null)
        //            builder.ForeColor = ConvertColor(parameters.ForegroundColor);

        //        if (parameters.BackgroundColor != null)
        //            builder.BackColor = ConvertColor(parameters.BackgroundColor);

        //        if (parameters.SymbolHeight != null)
        //        {
        //            builder.ImageHeight = ConvertSymbolHeight(parameters.SymbolHeight);
        //            builder.AutoSize = false;
        //        }

        //        builder.CodeLocation = CodeLocation.None;

        //        if (parameters.DisplayText)
        //            builder.CodeLocation = CodeLocation.Below;

        //        builder.CaptionAbove.Text = "";

        //        const float scale = 0.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
        //        float xdim = 1.0f;

        //        if (builder.SymbologyType == Symbology.QR)
        //        {
        //            builder.AutoSize = false;
        //            builder.ImageWidth *= scale;
        //            builder.ImageHeight = builder.ImageWidth;
        //            xdim = builder.ImageHeight / 25;
        //            builder.xDimension = builder.yDimension = xdim;
        //        }

        //        if (parameters.ScalingFactor != null)
        //        {
        //            float scalingFactor = ConvertScalingFactor(parameters.ScalingFactor);
        //            builder.ImageHeight *= scalingFactor;
        //            if (builder.SymbologyType == Symbology.QR)
        //            {
        //                builder.ImageWidth = builder.ImageHeight;
        //                builder.xDimension = builder.yDimension = xdim * scalingFactor;
        //            }

        //            builder.AutoSize = false;
        //        }

        //        return builder.BarCodeImage;
        //    }

        //    public Image GetOldBarcodeImage(BarcodeParameters parameters)
        //    {
        //        if (parameters.BarcodeType == null || parameters.BarcodeValue == null)
        //            return null;

        //        BarCodeBuilder builder = new BarCodeBuilder();

        //        builder.SymbologyType = ConvertBarcodeType(parameters.BarcodeType);
        //        if (builder.SymbologyType == (Symbology)int.MinValue)
        //            return null;

        //        builder.CodeText = parameters.BarcodeValue;

        //        if (builder.SymbologyType == Symbology.QR)
        //            builder.Display2DText = parameters.BarcodeValue;

        //        if (parameters.ForegroundColor != null)
        //            builder.ForeColor = ConvertColor(parameters.ForegroundColor);

        //        if (parameters.BackgroundColor != null)
        //            builder.BackColor = ConvertColor(parameters.BackgroundColor);

        //        if (parameters.SymbolHeight != null)
        //        {
        //            builder.ImageHeight = ConvertSymbolHeight(parameters.SymbolHeight);
        //            builder.AutoSize = false;
        //        }

        //        builder.CodeLocation = CodeLocation.None;

        //        if (parameters.DisplayText)
        //            builder.CodeLocation = CodeLocation.Below;

        //        builder.CaptionAbove.Text = "";

        //        const float scale = 0.4f; // Empiric scaling factor for converting Word barcode to Aspose.BarCode
        //        float xdim = 1.0f;

        //        if (builder.SymbologyType == Symbology.QR)
        //        {
        //            builder.AutoSize = false;
        //            builder.ImageWidth *= scale;
        //            builder.ImageHeight = builder.ImageWidth;
        //            xdim = builder.ImageHeight / 25;
        //            builder.xDimension = builder.yDimension = xdim;
        //        }

        //        if (parameters.ScalingFactor != null)
        //        {
        //            float scalingFactor = ConvertScalingFactor(parameters.ScalingFactor);
        //            builder.ImageHeight *= scalingFactor;
        //            if (builder.SymbologyType == Symbology.QR)
        //            {
        //                builder.ImageWidth = builder.ImageHeight;
        //                builder.xDimension = builder.yDimension = xdim * scalingFactor;
        //            }

        //            builder.AutoSize = false;
        //        }

        //        return builder.BarCodeImage;
        //    }
        //}

    }
}
