using Microsoft.Extensions.Configuration;
using SPOUtils;
using Utils;

namespace SPUserImageSync
{
    public class Config : CSOMConfig
    {
        public Config(IConfiguration config) : base(config)
        {
        }

        [ConfigValue]
        public bool SimulateSPOUpdatesOnly { get; set; }

        [ConfigValue]
        public string TenantName { get; set; } = string.Empty;

        [ConfigValue]
        public string? SPSecret { get; set; } = string.Empty;

        [ConfigValue]
        public string? SPClientID { get; set; } = string.Empty;

        [ConfigValue(true)]
        public string AppInsightsInstrumentationKey { get; set; } = string.Empty;

        public bool HaveAppInsightsConfigured => !string.IsNullOrEmpty(AppInsightsInstrumentationKey);
    }
}
