using Dockit.Convert;

if (args.Length != 4) return 2;
var lockPath = args[0];
var eventLog = args[1];
var holdMilliseconds = int.Parse(args[2], System.Globalization.CultureInfo.InvariantCulture);
var timeoutMilliseconds = int.Parse(args[3], System.Globalization.CultureInfo.InvariantCulture);
try
{
    using var lease = WpsRpcSession.AcquireSpreadsheetLease(TimeSpan.FromMilliseconds(timeoutMilliseconds), lockPath);
    File.AppendAllText(eventLog, $"+{Environment.ProcessId}\n");
    Thread.Sleep(holdMilliseconds);
    File.AppendAllText(eventLog, $"-{Environment.ProcessId}\n");
    return 0;
}
catch (TimeoutException)
{
    File.AppendAllText(eventLog, $"!{Environment.ProcessId}\n");
    return 23;
}
