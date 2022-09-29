using Microsoft.Graph;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.UserProfiles;
using SPOUtils;
using SPUserImageSync;

namespace ConsoleApp.Engine
{
    public class GraphUserProfileImageSyncroniser : BaseSyncClass
    {
        #region Constructor & Privates

        private ClientContext _adminCtx;
        private ClientContext _mySitesCtx;
        private PeopleManager _peopleManager;

        public GraphUserProfileImageSyncroniser(Config config, DebugTracer tracer, GraphServiceClient graphServiceClient, ClientContext adminCtx, ClientContext mySitesCtx, PeopleManager peopleManager) : base(config, tracer, graphServiceClient)
        {
            _adminCtx = adminCtx;
            _mySitesCtx = mySitesCtx;
            _peopleManager = peopleManager;
        }
        #endregion

        internal async Task SyncUser(string azureAdUsername)
        {
            var spProfileId = GetProfileId(azureAdUsername);

            var personProperties = _peopleManager.GetPropertiesFor(spProfileId);

            _adminCtx.Load(personProperties, p => p.AccountName, p => p.PictureUrl);
            try
            {
                await _adminCtx.ExecuteQueryAsyncWithThrottleRetries(_tracer);
            }
            catch (Exception ex)
            {
                _tracer.TrackException(ex);
                _tracer.TrackTrace($"Couldn't load SP properties for user '{azureAdUsername}'. Skipping for now. Error was:{ex.Message}");
                return;
            }

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
            catch (ServiceException ex)
            {
                _tracer.TrackTrace($"Error {ex.Message} loading profile for user '{azureAdUsername}'. Skipping user.");
                return;
            }

            var picRelativeUrl = $"/User Photos/Profile Pictures/{Guid.NewGuid()}_ExternalMigratedThumbnail.jpg";
            var fullPicUrl = $"{_mySitesCtx.Web.Url}{picRelativeUrl}";

            if (!_config.SimulateSPOUpdatesOnly)
            {
                // Get and upload Azure AD image to SPO library
                using (var graphImageStream = await _graphServiceClient.Users[user.Id].Photo.Content.Request().GetAsync())
                {
                    // Upload a file by adding it to the folder's files collection
                    var addedFile = siteAssetsFolder.Files.Add(new FileCreationInformation
                    {
                        Url = picRelativeUrl,
                        ContentStream = graphImageStream
                    });
                    await _mySitesCtx.ExecuteQueryAsyncWithThrottleRetries(_tracer);
                }
                _tracer.TrackTrace($"Uploaded profile picture to URL: {fullPicUrl}");

                // Update user profile
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
            else
            {
                _tracer.TrackTrace($"SIMULATION MODE: Uploaded profile picture to URL: {fullPicUrl}");
                _tracer.TrackTrace($"SIMULATION MODE: {azureAdUsername} profile updated succesfully for uploaded image from Azure AD");
            }
        }
    }
}
