using System;
using System.IO;

namespace TeXShift.Core.Utils
{
    /// <summary>
    /// Resolves the per-user directories TeXShift writes to at runtime.
    /// The uninstaller removes <see cref="LocalDataRoot"/> as a whole, so anything
    /// placed underneath it is cleaned up without extra installer work.
    /// </summary>
    public static class TeXShiftPaths
    {
        private const string RootFolderName = "TeXShift";
        private const string RuntimeLogFolderName = "logs";

        public static string LocalDataRoot => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            RootFolderName);

        /// <summary>
        /// Home of the always-on runtime log. Deliberately outside the debug output folder:
        /// runtime diagnostics are not a debug feature and must not follow the debug settings.
        /// </summary>
        public static string RuntimeLogFolder => Path.Combine(LocalDataRoot, RuntimeLogFolderName);
    }
}
