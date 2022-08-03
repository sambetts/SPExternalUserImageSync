
using ConsoleApp.Engine;
using SPOUtils;
using SPUserImageSync;


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

tracer.TrackTrace("Image Sync start-up");

var s = new AzureAdImageSyncer(config, tracer);
await s.FindAndSyncAllExternalUserImagesToSPO();

