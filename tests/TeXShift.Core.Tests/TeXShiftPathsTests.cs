using System;
using System.IO;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Logging;
using TeXShift.Core.Utils;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class TeXShiftPathsTests
    {
        [TestMethod]
        public void RuntimeLogFolderLivesUnderTheUninstalledDataRoot()
        {
            string localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);

            Assert.AreEqual(Path.Combine(localAppData, "TeXShift"), TeXShiftPaths.LocalDataRoot);
            Assert.AreEqual(Path.Combine(TeXShiftPaths.LocalDataRoot, "logs"), TeXShiftPaths.RuntimeLogFolder);
        }

        [TestMethod]
        public void RuntimeLogFolderIgnoresTheConfiguredDebugOutputPath()
        {
            // The runtime log is always-on diagnostics, so it must not follow the debug settings.
            // Pointing the debug output somewhere else must leave the runtime log where it is.
            string customDebugPath = Path.Combine(Path.GetTempPath(), "TeXShiftPathsTests");
            string debugFolder = DebugLogger.ResolveDebugOutputFolder(customDebugPath);

            Assert.AreEqual(customDebugPath, debugFolder);
            Assert.AreNotEqual(debugFolder, TeXShiftPaths.RuntimeLogFolder);
            StringAssert.EndsWith(TeXShiftPaths.RuntimeLogFolder, Path.Combine("TeXShift", "logs"));

            // Resolving must stay side-effect free.
            Assert.IsFalse(Directory.Exists(customDebugPath));
        }
    }
}
