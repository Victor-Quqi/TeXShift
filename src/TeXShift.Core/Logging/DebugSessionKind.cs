namespace TeXShift.Core.Logging
{
    /// <summary>
    /// Identifies which feature triggered a debug logging session.
    /// Used to disambiguate output filenames across forward/reverse conversion and debug tools.
    /// </summary>
    public enum DebugSessionKind
    {
        ForwardConversion,
        ReverseConversion,
        SelectionXml,
        PageXml
    }
}

