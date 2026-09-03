namespace Tiwater.Office;

internal static class WritableFileCopy
{
    public static void Copy(string source, string destination, bool overwrite = false)
    {
        var destinationExisted = File.Exists(destination);
        try
        {
            using var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var output = new FileStream(
                destination,
                overwrite ? FileMode.Create : FileMode.CreateNew,
                FileAccess.Write,
                FileShare.None);
            input.CopyTo(output);
        }
        catch
        {
            if (!destinationExisted && File.Exists(destination)) File.Delete(destination);
            throw;
        }
    }
}
