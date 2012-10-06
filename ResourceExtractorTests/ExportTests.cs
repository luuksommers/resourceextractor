using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ResourceExtractor;

namespace ResourceExtractorTests
{
    [TestClass]
    public class ExportTests
    {
        public TestContext TestContext { get; set; }

        [TestMethod]
        [DeploymentItem("Strings.nl.resx")]
        [DeploymentItem("Strings.resx")]
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
