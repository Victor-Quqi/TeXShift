namespace TeXShift.Core.OneNoteMeta
{
    internal static class TeXShiftMetaKeys
    {
        public const string Prefix = "texshift-";
        public const string Schema = "texshift-schema";
        public const string SchemaVersion = "1";

        public const string Mode = "texshift-mode";
        public const string ModeRender = "render";
        public const string ModeSource = "source";

        public const string SourceEncoding = "texshift-sourceEncoding";
        public const string EncodingPlainV1 = "plain-v1";
        public const string SourceChunkPrefix = "texshift-source-";

        public const string SigVersion = "texshift-sigVersion";
        public const string SigVersionValue = "1";
        public const string Sig = "texshift-sig";

        public const int MaxChunkLength = 8000;
    }
}
