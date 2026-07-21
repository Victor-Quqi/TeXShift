using System;
using System.Reflection;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OneNoteInterop = global::Microsoft.Office.Interop.OneNote;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class OneNoteInteropContractTests
    {
        [TestMethod]
        public void ApplicationContractUsesExpectedComIdentityAndDispIds()
        {
            var application = typeof(OneNoteInterop.IApplication);

            Assert.AreEqual(new Guid("452AC71A-B655-4967-A208-A4CC39DD7949"), application.GUID);
            var typeLibType = application.GetCustomAttribute<TypeLibTypeAttribute>().Value;
            Assert.IsTrue(typeLibType.HasFlag(TypeLibTypeFlags.FDual));
            Assert.IsTrue(typeLibType.HasFlag(TypeLibTypeFlags.FDispatchable));
            Assert.AreEqual(0x60020007, GetDispId(application.GetMethod("GetPageContent")));
            Assert.AreEqual(0x60020008, GetDispId(application.GetMethod("UpdatePageContent")));
            Assert.AreEqual(100, GetDispId(application.GetProperty("Windows").GetMethod));
        }

        private static int GetDispId(MemberInfo member)
        {
            return member.GetCustomAttribute<DispIdAttribute>().Value;
        }
    }
}
