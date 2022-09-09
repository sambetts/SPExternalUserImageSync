using Azure.Identity;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using PnP.Framework;
using SPOUtils;
using SPUserImageSync;

namespace ConsoleApp.Engine
{
    /// <summary>
    /// Main app logic entry-point
    /// </summary>
    public class AzureAdImageSyncer : BaseSyncClass, IDisposable
    {
        #region Constructor & Privates

        private ClientContext _adminCtx;
        private ClientContext _mySitesCtx;
        private PeopleManager _peopleManager;

        public AzureAdImageSyncer(Config config, DebugTracer tracer) : base(config, tracer, GetGraphClient(config))
        {

            var siteUrlAdmin = $"https://{_config.TenantName}-admin.sharepoint.com";
            var siteUrlMySite = $"https://{_config.TenantName}-my.sharepoint.com";
            _adminCtx = new AuthenticationManager().GetACSAppOnlyContext(siteUrlAdmin, _config.SPClientID, _config.SPSecret);
            _mySitesCtx = new AuthenticationManager().GetACSAppOnlyContext(siteUrlMySite, _config.SPClientID, _config.SPSecret);
            _peopleManager = new PeopleManager(_adminCtx);
        }

        static GraphServiceClient GetGraphClient(Config config)
        {
            var c = new GraphServiceClient(
                new ClientSecretCredential(config.AzureAdConfig.TenantId, config.AzureAdConfig.ClientID, config.AzureAdConfig.Secret));
            c.HttpProvider.OverallTimeout = TimeSpan.FromMinutes(60);
            return c;
        }

        #endregion

        public async Task FindAndSyncAllExternalUserImagesToSPO()
        {
            // Read external users
            List<Microsoft.Graph.User> allUsers;
            try
            {
                allUsers = await Get(_graphServiceClient.Users.Request().Filter("userType eq 'Guest'"));
            }
            catch (ServiceException ex)
            {
                _tracer.TrackTrace($"Couldn't read Graph users: {ex.Message}");
                return;
            }
            if (allUsers.Count == 0)
            {
                _tracer.TrackTrace($"Found {allUsers.Count} external users. Nothing to do.");
                return;
            }
            else
            {
                _tracer.TrackTrace($"Found {allUsers.Count} external users.");
            }

            // Test access
            _tracer.TrackTrace("Testing access to SharePoint...");
            var spProfileId = GetProfileId(allUsers[0].UserPrincipalName);
            try
            {
                var personProperties = _peopleManager.GetPropertiesFor(spProfileId);

                _mySitesCtx.Load(_mySitesCtx.Web, w => w.Title, w => w.Url, w => w.Folders);
                await _mySitesCtx.ExecuteQueryAsyncWithThrottleRetries(_tracer);

                _adminCtx.Load(personProperties, p => p.AccountName, p => p.UserProfileProperties);
                await _adminCtx.ExecuteQueryAsyncWithThrottleRetries(_tracer);
            }
            catch (Exception ex)
            {
                _tracer.TrackTrace($"Can't access SharePoint - got error {ex.Message}");
                return;
            }

            // Start sync for each user
            var s = new GraphUserProfileImageSyncroniser(_config, _tracer, _graphServiceClient, _adminCtx, _mySitesCtx, _peopleManager);
            foreach (var externalUser in allUsers)
            {
                await s.SyncUser(externalUser.UserPrincipalName);
            }
        }

        private async Task<List<Microsoft.Graph.User>> Get(IGraphServiceUsersCollectionRequest request)
        {
            var allUsers = new List<Microsoft.Graph.User>();
            var nextRequest = request;
            while (nextRequest != null)
            {
                var results = await nextRequest.GetAsync();
                nextRequest = results.NextPageRequest;
                allUsers.AddRange(results);
            }

            return allUsers;
        }
        public void Dispose()
        {
            _adminCtx.Dispose();
            _mySitesCtx.Dispose();
        }

    }
}
