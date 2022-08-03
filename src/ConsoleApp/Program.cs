
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

if (config.SimulateSPOUpdatesOnly)
{
    tracer.TrackTrace("Image Sync start-up - SIMULATION ONLY");
}
else
{
    tracer.TrackTrace("Image Sync start-up - SPO WRITES ENABLED IN CONFIG - Simulation only mode is disabled!");
#if DEBUG
    Console.WriteLine("Waiting 5 seconds in case you change your mind...");
    Thread.Sleep(5000);
#endif
}
Console.WriteLine($"We will look for all external users in tenant '{config.AzureAdConfig.TenantId}' and set any SharePoint Online user profile with the same image in Azure AD, if one is not already set.");

var s = new AzureAdImageSyncer(config, tracer);
await s.FindAndSyncAllExternalUserImagesToSPO();

