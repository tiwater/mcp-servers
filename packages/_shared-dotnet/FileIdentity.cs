using System.Security.Cryptography;

namespace Tiwater.RuntimeContracts;

public static class FileIdentity
{
    public static FileContentIdentity IdentifyFile(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
        {
            throw new ArgumentException("File path must be non-empty.", nameof(filePath));
        }

        var resolvedPath = Path.GetFullPath(filePath);
        using var stream = new FileStream(resolvedPath, FileMode.Open, FileAccess.Read, FileShare.Read);
        var sizeBytes = stream.Length;
        var digest = SHA256.HashData(stream);
        return Create(resolvedPath, sizeBytes, digest);
    }

    public static FileContentIdentity IdentifyBytes(string path, ReadOnlySpan<byte> bytes)
    {
        if (string.IsNullOrWhiteSpace(path))
        {
            throw new ArgumentException("Path must be non-empty.", nameof(path));
        }

        return Create(path, bytes.Length, SHA256.HashData(bytes));
    }

    internal static ArtifactIdentity IdentifyArtifactBytes(
        ReadOnlySpan<byte> bytes,
        string mediaType,
        string encoding,
        SchemaIdentity schema)
    {
        var digest = SHA256.HashData(bytes);
        var sha256 = Convert.ToHexStringLower(digest);
        return new ArtifactIdentity(
            ArtifactId: $"sha256:{sha256}",
            SizeBytes: bytes.Length,
            Sha256: sha256,
            MediaType: mediaType,
            Encoding: encoding,
            Schema: schema);
    }

    private static FileContentIdentity Create(string path, long sizeBytes, byte[] digest)
    {
        var sha256 = Convert.ToHexStringLower(digest);
        return new FileContentIdentity(path, sizeBytes, sha256, $"sha256:{sha256}");
    }
}
