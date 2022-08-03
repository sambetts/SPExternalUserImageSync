using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Utils;

namespace SPOUtils
{
    public class CSOMConfig : BaseConfig
    {
        public CSOMConfig(IConfiguration config) : base(config)
        {
        }

        [ConfigSection("AzureAd")]
        public AzureAdConfig AzureAdConfig { get; set; } = null!;
    }
}
