using System;
using System.IO;
using System.Text;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ResourceExtractor;

namespace ResourceExtractorTests
{
    [TestClass]
    public class ExportTests
    {
        public TestContext TestContext { get; set; }

        [TestMethod]
        public void Export()
        {
            // Assign
            var outputPath = Path.Combine(TestContext.DeploymentDirectory, "Test.xls");
            if( File.Exists(outputPath))
            {
                File.Delete(outputPath);
            }

            var exporter = new ResourceExport(TestContext.DeploymentDirectory, "en", outputPath, false);

            // Act
            exporter.Export();

            // Assert
            Assert.IsTrue(File.Exists(outputPath));
        }
    }
}
