using Dockit.Convert;

if (args.Length != 5) return 2;
var route = args[0];
var lockPath = args[1];
var eventLog = args[2];
var holdMilliseconds = int.Parse(args[3], System.Globalization.CultureInfo.InvariantCulture);
var timeoutMilliseconds = int.Parse(args[4], System.Globalization.CultureInfo.InvariantCulture);
try
{
    using var lease = route switch
    {
        "writer" => WpsPdfConverter.AcquireRuntimeLease(TimeSpan.FromMilliseconds(timeoutMilliseconds), lockPath),
        "spreadsheet" => EtPdfConverter.AcquireRuntimeLease(TimeSpan.FromMilliseconds(timeoutMilliseconds), lockPath),
        "presentation" => WppPdfConverter.AcquireRuntimeLease(TimeSpan.FromMilliseconds(timeoutMilliseconds), lockPath),
        "lima" => LimaWpsPdfConverter.AcquireOfficeHostLease(TimeSpan.FromMilliseconds(timeoutMilliseconds), lockPath),
        _ => throw new ArgumentOutOfRangeException(nameof(route), route, "Unknown Office route."),
    };
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
