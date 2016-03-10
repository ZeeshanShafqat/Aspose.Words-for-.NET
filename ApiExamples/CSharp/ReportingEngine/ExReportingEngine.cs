// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System;
using System.IO;
using System.Data;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Reporting;
using NUnit.Framework;

namespace ApiExamples
{
    [TestFixture]
    public class ExReportingEngine : ApiExampleBase
    {
        private readonly string image = MyDir + @"Images\Test_636_852.gif";

        [Test]
        public void StretchImageFitHeight()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<image [src.Image] -fitHeight>>");

            ImageStream imageStream = new ImageStream(new FileStream(this.image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);

                //Assert that width is keeped and height is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        public void StretchImageFitWidth()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<image [src.Image] -fitWidth>>");

            ImageStream imageStream = new ImageStream(new FileStream(this.image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox and 
                Assert.IsTrue(shape.ImageData.HasImage);

                //Assert that height is keeped and width is changed
                Assert.AreNotEqual(431.5, shape.Width);
                Assert.AreEqual(346.35, shape.Height);
            }

            dstStream.Dispose();
        }

        [Test]
        public void StretchImageFitSize()
        {
            Document doc = DocumentHelper.CreateTemplateDocumentForReportingEngine("<<image [src.Image] -fitSize>>");

            ImageStream imageStream = new ImageStream(new FileStream(this.image, FileMode.Open, FileAccess.Read));

            BuildReport(doc, imageStream, "src");

            MemoryStream dstStream = new MemoryStream();
            doc.Save(dstStream, SaveFormat.Docx);

            doc = new Document(dstStream);

            NodeCollection shapes = doc.GetChildNodes(NodeType.Shape, true);

            foreach (Shape shape in shapes)
            {
                // Assert that the image is really insert in textbox 
                Assert.IsTrue(shape.ImageData.HasImage);

                //Assert that height is changed and width is changed
                Assert.AreNotEqual(346.35, shape.Height);
                Assert.AreNotEqual(431.5, shape.Width);
            }

            dstStream.Dispose();
        }

        [Test]
        [ExpectedException(typeof(InvalidOperationException))]
        public void AllowMissingDataFieldsException()
        {
            Document doc = new Document();

            DocumentHelper.InsertNewRun(doc, "<<if [value == “true”] >>ok<<else>>Cancel<</if>>");

            DataSet dataSet = new DataSet();
            dataSet.ReadXml(MyDir + "DataSet.xml", XmlReadMode.InferSchema);

            BuildReport(doc, dataSet, "Bad");
        }

        /// <summary>
        /// Assert that the exception from previous test is not repeated with AllowMissingMembers parameter
        /// </summary>
        [Test]
        public void AllowMissingDataFields()
        {
            Document doc = new Document();

            DocumentHelper.InsertNewRun(doc, "<<if [value == “true”] >>ok<<else>>Cancel<</if>>");

            DataSet dataSet = new DataSet();
            dataSet.ReadXml(MyDir + "DataSet.xml", XmlReadMode.InferSchema);

            ReportingEngine engine = new ReportingEngine();
            engine.Options = ReportBuildOptions.AllowMissingMembers;

            engine.BuildReport(doc, dataSet, "Bad");
        }

        private static void BuildReport(Document document, object dataSource, string dataSourceName)
        {
            ReportingEngine engine = new ReportingEngine();
            engine.BuildReport(document, dataSource, dataSourceName);
        }
    }
}

public class ImageStream
{
    public ImageStream(Stream stream)
    {
        this.Image = stream;
    }

    public Stream Image { get; set; }
}