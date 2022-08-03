using Microsoft.Graph;
using SPOUtils;
using SPUserImageSync;

namespace ConsoleApp.Engine
{

    public abstract class BaseSyncClass
    {
        protected readonly Config _config;
        protected readonly DebugTracer _tracer;
        protected readonly GraphServiceClient _graphServiceClient;
        public BaseSyncClass(Config config, DebugTracer tracer, GraphServiceClient graphServiceClient)
        {
            this._config = config;
            this._tracer = tracer;

            this._graphServiceClient = graphServiceClient;

        }
        protected string GetProfileId(string azureAdUsername)
        {
            return $"i:0#.f|membership|{azureAdUsername}";
        }
    }
}
