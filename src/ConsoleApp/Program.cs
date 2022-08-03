// See https://aka.ms/new-console-template for more information
using SPOUtils;
using SPUserImageSync;

Console.WriteLine("Image Sync start-up");

var config = ConsoleUtils.GetConfigurationWithDefaultBuilder();
ConsoleUtils.PrintCommonStartupDetails();

// Send to application insights or just the stdout?
DebugTracer tracer;
if (config.HaveAppInsightsConfigured)
{
    tracer = new DebugTracer(config.AppInsightsInstrumentationKey, "Indexer");
}
else
    tracer = DebugTracer.ConsoleOnlyTracer();


var s = new ImgSync(config, tracer);
await s.Go("sambetts_microsoft.com#EXT#@M365x72460609.onmicrosoft.com");

