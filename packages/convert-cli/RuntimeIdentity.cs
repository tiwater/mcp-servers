namespace Dockit.Convert;

internal static class RuntimeIdentity
{
    public static string Version
    {
        get
        {
            var version = typeof(RuntimeIdentity).Assembly.GetName().Version
                ?? throw new InvalidOperationException("Convert runtime assembly version is unavailable.");
            return $"{version.Major}.{version.Minor}.{version.Build}";
        }
    }
}
