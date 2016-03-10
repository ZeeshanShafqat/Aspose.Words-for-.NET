using Aspose.Words;
using Aspose.Words.Fields;

using NUnit.Framework;

namespace ApiExamples
{
    using System;

    [TestFixture]
    internal class ExParagraph : ApiExampleBase
    {
        [Test]
        public void InsertField()
        {
            //ExStart
            //ExFor:Paragraph.InsertField
            //ExSummary:Shows how to insert field using several methods: "field code", "field code and field value", "field code and field value after a run of text"
            Aspose.Words.Document doc = new Aspose.Words.Document();

            //Get the first paragraph of the document
            Paragraph para = doc.FirstSection.Body.FirstParagraph;

            //Inseting field using field code
            //Note: All methods support inserting field after some node. Just set "true" in the "isAfter" parameter
            para.InsertField(" AUTHOR ", null, false);

            //Using field type
            //Note:
            //1. For inserting field using field type, you can choose, update field before or after you open the document ("updateField" parameter)
            //2. For other methods it's works automatically
            para.InsertField(FieldType.FieldAuthor, false, null, true);

            //Using field code and field value
            para.InsertField(" AUTHOR ", "Test Field Value", null, false);

            //Add a run of text
            Run run = new Run(doc) { Text = " Hello World!" };
            para.AppendChild(run);

            //Using field code and field value before a run of text
            //Note: For inserting field before/after a run of text you can use all methods above, just add ref on your text ("refNode" parameter)
            para.InsertField(" AUTHOR ", "Test Field Value", run, false);
            //ExEnd
        }

        [Test]
        public void InsertFieldBeforeText()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " AUTHOR ", null, false);

            Assert.AreEqual("\u0013 AUTHOR \u0014Test Author\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldAfterText()
        {
            string date = DateTime.Today.ToString("d");

            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCode(doc, " DATE ", null, true);

            Assert.AreEqual(String.Format("Hello World!\u0013 DATE \u0014{0}\u0015\r", date), DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldBeforeTextWithoutUpdateField()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, false);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldAfterTextWithoutUpdateField()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, true);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldWithoutSeparator()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldListNum, true, null, false);

            Assert.AreEqual("\u0013 LISTNUM \u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldBeforeParagraphWithoutDocumentAuthor()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();
            doc.BuiltInDocumentProperties.Author = "";

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, false);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldAfterParagraphWithoutChangingDocumentAuthor()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", null, null, true);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldBeforeRunText()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            //Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!");

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "Test Field Value", run, false);

            Assert.AreEqual("Hello World!\u0013 AUTHOR \u0014Test Field Value\u0015 Hello World!\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        [Test]
        public void InsertFieldAfterRunText()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentFillWithDummyText();

            //Add some text into the paragraph
            Run run = DocumentHelper.InsertNewRun(doc, " Hello World!");

            InsertFieldUsingFieldCodeFieldString(doc, " AUTHOR ", "", run, true);

            Assert.AreEqual("Hello World! Hello World!\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        /// <summary>
        /// Test for WORDSNET-12396
        /// </summary>
        [Test]
        public void InsertFieldEmptyParagraphWithoutUpdateField()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentWithoutDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, false, null, false);

            Assert.AreEqual("\u0013 AUTHOR \u0014\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        /// <summary>
        /// Test for WORDSNET-12397
        /// </summary>
        [Test]
        public void InsertFieldEmptyParagraphWithUpdateField()
        {
            Aspose.Words.Document doc = DocumentHelper.CreateDocumentWithoutDummyText();

            InsertFieldUsingFieldType(doc, FieldType.FieldAuthor, true, null, false);

            Assert.AreEqual("\u0013 AUTHOR \u0014Test Author\u0015\r", DocumentHelper.GetParagraphText(doc, 0));
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field type
        /// </summary>
        private static void InsertFieldUsingFieldType(Aspose.Words.Document doc, FieldType fieldType, bool updateField, Aspose.Words.Node refNode, bool isAfter)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, 0);
            para.InsertField(fieldType, updateField, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code
        /// </summary>
        private static void InsertFieldUsingFieldCode(Aspose.Words.Document doc, string fieldCode, Aspose.Words.Node refNode, bool isAfter)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, 0);
            para.InsertField(fieldCode, refNode, isAfter);
        }

        /// <summary>
        /// Insert field into the first paragraph of the current document using field code and field string
        /// </summary>
        private static void InsertFieldUsingFieldCodeFieldString(Aspose.Words.Document doc, string fieldCode, string fieldValue, Aspose.Words.Node refNode, bool isAfter)
        {
            Paragraph para = DocumentHelper.GetParagraph(doc, 0);
            para.InsertField(fieldCode, fieldValue, refNode, isAfter);
        }
    }
}
