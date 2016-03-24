// Copyright (c) 2001-2016 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Words. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////

using System.Drawing.Imaging;
using NUnit.Framework;

namespace ApiExamples
{
    using Aspose.Words;
    using Aspose.Words.Saving;

    [TestFixture]
    internal class ExImageSaveOptions : ApiExampleBase
    {
        //Todo: add as example
        [Test]
        public void UseGdiEmfRenderer()
        {
            Document doc = new Document(MyDir + "MyraidPro Sample.docx");

            ImageSaveOptions saveOptions = new ImageSaveOptions(SaveFormat.Emf);
            saveOptions.UseGdiEmfRenderer = false;
        }
    }
}
