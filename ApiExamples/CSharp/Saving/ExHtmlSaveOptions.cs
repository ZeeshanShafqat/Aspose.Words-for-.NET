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
    internal class ExHtmlSaveOptions : ApiExampleBase
    {
        //For assert this test you need to open documents and html do not have negative left margins
        [Test]
        [TestCase(SaveFormat.Html)]
        [TestCase(SaveFormat.Mhtml)]
        [TestCase(SaveFormat.Epub)]
        public void ExportPageMargins(SaveFormat saveFormat)
        {
            Document doc = new Document(MyDir + "HtmlSaveOptions.ExportPageMargins.docx");

            HtmlSaveOptions htmlSaveOptions = new HtmlSaveOptions
            {
                SaveFormat = saveFormat, 
                ExportPageMargins = true
            };

            switch (saveFormat)
            {
                case SaveFormat.Html:
                    doc.Save(MyDir + "ExportPageMargins.html", htmlSaveOptions);
                    break;
                case SaveFormat.Mhtml:
                    doc.Save(MyDir + "ExportPageMargins.Mhtml", htmlSaveOptions);
                    break;
                case SaveFormat.Epub:
                    doc.Save(MyDir + "ExportPageMargins.Epub", htmlSaveOptions); //There is draw images bug with epub. Need to write NSezganov
                    break;
            }
        }
    }
}
