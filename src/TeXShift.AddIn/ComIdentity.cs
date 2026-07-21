namespace TeXShift.AddIn
{
    internal static class ComIdentity
    {
#if DEBUG
        public const string Clsid = "6AF190C5-74D6-4692-A053-C9EB6E8B22D1";
        public const string ProgId = "TeXShift.AddIn.Debug.Connect";
#else
        public const string Clsid = "1EE8F914-ECBD-4709-92C0-E770851C4966";
        public const string ProgId = "TeXShift.AddIn.Connect";
#endif
    }
}
