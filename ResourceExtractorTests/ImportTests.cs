using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using ResourceExtractor;

namespace ResourceExtractorTests
{
    [TestClass]
    public class ImportTests
    {
        public TestContext TestContext { get; set; }


        [TestMethod]
        [DeploymentItem("Import.xls")]
        public void Import()
        {
            // Assign
            var importPath = Path.Combine(TestContext.DeploymentDirectory, "Import.xls");

            var exporter = new ResourceImport(importPath, "en", TestContext.DeploymentDirectory);

            // Act
            exporter.Import();

            // Assert
            Assert.IsTrue(File.Exists(Path.Combine(TestContext.DeploymentDirectory, "Strings.resx")));
            Assert.IsTrue(File.Exists(Path.Combine(TestContext.DeploymentDirectory, "Strings.nl.resx")));
        }
    }   
}
