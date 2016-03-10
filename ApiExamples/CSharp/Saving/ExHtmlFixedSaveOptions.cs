// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Text;

using Aspose.Words;
using Aspose.Words.Saving;

using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    internal class ExHtmlFixedSaveOptions : ApiExampleBase
    {
        [Test]
        public void UseEncoding()
        {
            //ExStart
            //ExFor:Saving.HtmlFixedSaveOptions.Encoding
            //ExSummary:Shows how to use "Encoding" parameter with "HtmlFixedSaveOptions"
            Document doc = new Document();

            DocumentBuilder builder = new DocumentBuilder(doc);
            builder.Writeln("Hello World!");

            //Create "HtmlFixedSaveOptions" with "Encoding" parameter
            //You can also set "Encoding" using System.Text.Encoding, like "Encoding.ASCII", or "Encoding.GetEncoding()"
            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = new ASCIIEncoding(),
                SaveFormat = SaveFormat.HtmlFixed,
            };

            //Uses "HtmlFixedSaveOptions"
            doc.Save(MyDir + "UseEncoding.html", htmlFixedSaveOptions);
            //ExEnd
        }

        //ToDo: maybe unit tests with "UseEncoding"
        //Note: Tests doesn't containt validation result, because it's may take a lot of time for assert result
        //For validation result, you can save the document to html file and check out with notepad++, that file encoding will be correctly displayed (Encoding tab in Notepad++)
        [Test]
        public void EncodingUsingSystemTextEncoding()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = Encoding.ASCII,
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedCss = true,
                ExportEmbeddedFonts = true,
                ExportEmbeddedImages = true,
                ExportEmbeddedSvg = true
            };

            doc.Save(MyDir + "EncodingUsingSystemTextEncoding_OUT.html", htmlFixedSaveOptions);
        }

        [Test]
        public void EncodingUsingNewEncoding()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = new UTF32Encoding(),
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedCss = true,
                ExportEmbeddedFonts = true,
                ExportEmbeddedImages = true,
                ExportEmbeddedSvg = true
            };

            doc.Save(MyDir + "EncodingUsingNewEncoding_OUT.html", htmlFixedSaveOptions);
        }

        [Test]
        public void EncodingUsingGetEncoding()
        {
            Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            HtmlFixedSaveOptions htmlFixedSaveOptions = new HtmlFixedSaveOptions
            {
                Encoding = Encoding.GetEncoding("utf-16"),
                SaveFormat = SaveFormat.HtmlFixed,
                ExportEmbeddedCss = true,
                ExportEmbeddedFonts = true,
                ExportEmbeddedImages = true,
                ExportEmbeddedSvg = true
            };

            doc.Save(MyDir + "EncodingUsingGetEncoding_OUT.html", htmlFixedSaveOptions);
        }
    }
}
