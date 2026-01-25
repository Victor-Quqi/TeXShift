using System;
using System.Linq;
using System.Xml.Linq;
using TeXShift.Core.OneNoteToMarkdown.Abstractions;

namespace TeXShift.Core.OneNoteToMarkdown.Handlers
{
    internal static class OneNoteTableHelpers
    {
        internal static XElement GetTable(XElement element, IOneNoteConverterContext context)
        {
            if (element == null)
            {
                return null;
            }

            var ns = context.OneNoteNamespace;
            return element.Element(ns + "Table");
        }

        internal static bool IsTrue(string value)
        {
            return string.Equals(value, "true", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool IsFalse(string value)
        {
            return string.Equals(value, "false", StringComparison.OrdinalIgnoreCase);
        }

        internal static bool TryGetSingleCell(XElement table, XNamespace ns, out XElement cell)
        {
            cell = null;
            if (table == null)
            {
                return false;
            }

            var rows = table.Elements(ns + "Row").ToList();
            if (rows.Count != 1)
            {
                return false;
            }

            var cells = rows[0].Elements(ns + "Cell").ToList();
            if (cells.Count != 1)
            {
                return false;
            }

            cell = cells[0];
            return true;
        }

        internal static bool IsLockedSingleColumn(XElement table, XNamespace ns)
        {
            if (table == null)
            {
                return false;
            }

            var column = table.Element(ns + "Columns")?.Elements(ns + "Column").FirstOrDefault();
            var isLocked = column?.Attribute("isLocked")?.Value;
            return string.Equals(isLocked, "true", StringComparison.OrdinalIgnoreCase);
        }
    }
}

