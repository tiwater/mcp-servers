using System.Security.Cryptography;
using System.Text;

namespace Tiwater.Office;

internal sealed class OutputWriteLease : IDisposable
{
    private readonly List<FileStream> handles = [];

    public static OutputWriteLease Acquire(params string[] paths)
    {
        var lease = new OutputWriteLease();
        try
        {
            var root = Path.Combine(Path.GetTempPath(), "tiwater-output-write-locks");
            if (OperatingSystem.IsWindows()) Directory.CreateDirectory(root);
            else Directory.CreateDirectory(root, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
            foreach (var identity in paths.Select(CanonicalPath).Select(value =>
                OperatingSystem.IsWindows() || OperatingSystem.IsMacOS() ? value.ToUpperInvariant() : value)
                .Distinct(StringComparer.Ordinal).Order(StringComparer.Ordinal))
            {
                var key = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(identity.Normalize())));
                var options = new FileStreamOptions { Mode = FileMode.OpenOrCreate, Access = FileAccess.ReadWrite, Share = FileShare.None };
                if (!OperatingSystem.IsWindows()) options.UnixCreateMode = UnixFileMode.UserRead | UnixFileMode.UserWrite;
                try { lease.handles.Add(new FileStream(Path.Combine(root, key + ".lock"), options)); }
                catch (IOException error) { throw new IOException("output-write-conflict-or-lock-unavailable", error); }
            }
            return lease;
        }
        catch { lease.Dispose(); throw; }
    }

    private static string CanonicalPath(string path)
    {
        var full = Path.GetFullPath(path);
        var parent = Path.GetDirectoryName(full);
        if (parent is null) return full;
        var resolved = Path.Combine(CanonicalPath(parent), Path.GetFileName(full));
        FileSystemInfo info = Directory.Exists(resolved) ? new DirectoryInfo(resolved) : new FileInfo(resolved);
        return info.LinkTarget is null ? resolved : info.ResolveLinkTarget(true)!.FullName;
    }

    public void Dispose()
    {
        foreach (var handle in handles) handle.Dispose();
        handles.Clear();
        // Never unlink guard files: waiters must all address the same inode.
    }
}
