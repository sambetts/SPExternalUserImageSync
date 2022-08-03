using Microsoft.ApplicationInsights;
using Microsoft.ApplicationInsights.Extensibility;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SPOUtils
{

    /// <summary>
    /// Unified console & AppInsights tracer
    /// </summary>
    public class DebugTracer
    {
        private TelemetryClient? AppInsights { get; set; }

        #region Constructors

        private DebugTracer() : this(string.Empty, string.Empty)
        {
        }
        public DebugTracer(string appInsightsKey, string context)
        {
            if (!string.IsNullOrEmpty(appInsightsKey))
            {
                Console.WriteLine($"Telemetry: sending runtime data to Application Insights with instrumentation key '{appInsightsKey}'");
                var configuration = TelemetryConfiguration.CreateDefault();
                configuration.InstrumentationKey = appInsightsKey;

                this.AppInsights = new TelemetryClient(configuration); ;
            }
            if (!string.IsNullOrEmpty(context) && AppInsights != null)
            {
                AppInsights.Context.Operation.Name = context;
            }
        }

        public static DebugTracer ConsoleOnlyTracer() { return new DebugTracer(); }


        #endregion

        public void TrackException(Exception ex)
        {
            if (AppInsights != null)
            {
                AppInsights.TrackException(ex);
            }
        }

        public void TrackTrace(string sayWut, Microsoft.ApplicationInsights.DataContracts.SeverityLevel severityLevel)
        {
            if (severityLevel != Microsoft.ApplicationInsights.DataContracts.SeverityLevel.Verbose)
            {
                Console.WriteLine($"[{DateTime.Now.ToString("G")}] {sayWut}");
            }

            if (AppInsights != null)
            {
                AppInsights.TrackTrace(sayWut, severityLevel);
            }
        }

        public void TrackTrace(string sayWut)
        {
            TrackTrace(sayWut, Microsoft.ApplicationInsights.DataContracts.SeverityLevel.Information);
        }
    }
}
