using Azure.Identity;
using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using PnP.Framework;
using SPOUtils;

namespace SPUserImageSync
{
    internal class ImgSync : IDisposable
    {
        #region Constructor & Privates

        private Config _config;
        private DebugTracer _tracer;
        private ClientContext _adminCtx;
        private ClientContext _mySitesCtx;
        private PeopleManager _peopleManager;
        private readonly GraphServiceClient _graphServiceClient;

        public ImgSync(Config config, DebugTracer tracer)
        {
            this._config = config;
            this._tracer = tracer;
            var siteUrlAdmin = $"https://{_config.SPOHostname}-admin.sharepoint.com";
            var siteUrlMySite = $"https://{_config.SPOHostname}-my.sharepoint.com";
            _adminCtx = new AuthenticationManager().GetACSAppOnlyContext(siteUrlAdmin, _config.SPClientID, _config.SPSecret);
            _mySitesCtx = new AuthenticationManager().GetACSAppOnlyContext(siteUrlMySite, _config.SPClientID, _config.SPSecret);
            _peopleManager = new PeopleManager(_adminCtx);
            this._graphServiceClient = new GraphServiceClient(
                new ClientSecretCredential(_config.AzureAdConfig.TenantId, _config.AzureAdConfig.ClientID, _config.AzureAdConfig.Secret));

        }
        #endregion

        internal async Task Go(string azureAdUsername)
        {
            var spProfileId = GetProfileId(azureAdUsername);
            var personProperties = _peopleManager.GetPropertiesFor(spProfileId);

            // Test access
            _tracer.TrackTrace("Testing access to SharePoint...");
            try
            {
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

            await VerifyUserProfile(azureAdUsername);
        }


        private async Task VerifyUserProfile(string azureAdUsername)
        {
            var spProfileId = GetProfileId(azureAdUsername);

            var personProperties = _peopleManager.GetPropertiesFor(spProfileId);

            _adminCtx.Load(personProperties, p => p.AccountName, p => p.PictureUrl);
            await _adminCtx.ExecuteQueryAsyncWithThrottleRetries(_tracer);

            var url = string.Empty;
            try
            {
                url = personProperties.PictureUrl;
            }
            catch (ServerObjectNullReferenceException)
            {
                // Ignore
            }
            catch (PropertyOrFieldNotInitializedException)
            {
                // Ignore
            }
            if (string.IsNullOrEmpty(url))
            {
                _tracer.TrackTrace($"{azureAdUsername} has no profile image URL in SPO");
                await SetImageWithAzureAdImage(azureAdUsername);
            }
            else
            {
                _tracer.TrackTrace($"{azureAdUsername} has profile image URL '{url}' in SPO");
            }
        }

        private async Task SetImageWithAzureAdImage(string azureAdUsername)
        {
            // Get a reference to a folder
            var siteAssetsFolder = _mySitesCtx.Web.Folders.Where(f => f.Name == "User Photos").FirstOrDefault();

            if (siteAssetsFolder == null)
            {
                _tracer.TrackTrace($"Can't find 'User Photos' list...");
                return;
            }

            var encodedUsername = System.Net.WebUtility.UrlEncode(azureAdUsername);
            Microsoft.Graph.User user;
            try
            {
                user = await _graphServiceClient.Users[encodedUsername].Request().GetAsync();
            }
            catch (ServiceException ex) when (ex.Message.Contains("Request_ResourceNotFound"))
            {
                _tracer.TrackTrace($"User '{azureAdUsername}' not found in Graph");
                return;
            }

            ProfilePhoto photoInfo;
            try
            {
                photoInfo = await _graphServiceClient.Users[user.Id].Photo.Request().GetAsync();
            }
            catch (ServiceException ex) when (ex.Message.Contains("ImageNotFound"))
            {
                _tracer.TrackTrace($"User '{azureAdUsername}' has no photo in Azure AD. Skipping user.");
                return;
            }

            var picRelativeUrl = $"/User Photos/Profile Pictures/{Guid.NewGuid()}_ExternalMigratedThumbnail.jpg";
            var fullPicUrl = $"{_mySitesCtx.Web.Url}{picRelativeUrl}";


            using (var graphImageStream = await _graphServiceClient.Users[user.Id].Photo.Content.Request().GetAsync())
            {
                // Upload a file by adding it to the folder's files collection
                var addedFile = siteAssetsFolder.Files.Add(new FileCreationInformation 
                { 
                    Url = picRelativeUrl, ContentStream = graphImageStream 
                });
                await _mySitesCtx.ExecuteQueryAsyncWithThrottleRetries(_tracer);
            }
            _tracer.TrackTrace($"Uploaded {fullPicUrl}");


            _peopleManager.SetSingleValueProfileProperty(GetProfileId(azureAdUsername), "PictureURL", fullPicUrl);

            try
            {
                await _adminCtx.ExecuteQueryAsyncWithThrottleRetries(_tracer);
                _tracer.TrackTrace($"{azureAdUsername} profile updated succesfully for uploaded image from Azure AD");
            }
            catch (ServerException ex)
            {
                // May get "User Profile Error 1000: User Not Found: Could not load profile data from the database."
                _tracer.TrackTrace($"{azureAdUsername} profile update failed - {ex.Message}");
            }
        }

        string GetProfileId(string azureAdUsername)
        {
            return $"i:0#.f|membership|{azureAdUsername}";
        }
        public void Dispose()
        {
            _adminCtx.Dispose();
            _mySitesCtx.Dispose();
        }
    }
}
