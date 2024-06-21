using Microsoft.Identity.Client;

namespace MauiSyncOutlookCalendar
{
    public partial class App : Application
    {
        internal static IPublicClientApplication ClientApplication;
        private string clientID, tenantID, authority;
        public App()
        {
            InitializeComponent();

            MainPage = new MainPage();

            //// You need to replace your Application or Client ID
            clientID = "09893c5e-c8e6-4652-9e11-43baa5422854";

            //// You need to replace your tenant ID
            tenantID = "77f1fe12-b049-4919-8c50-9fb41e5bb63b";

            authority = "https://login.microsoftonline.com/" + tenantID;

            ClientApplication = PublicClientApplicationBuilder.Create(clientID)
            .WithAuthority(authority)
            .WithRedirectUri("https://login.microsoftonline.com/common/oauth2/nativeclient")
            .Build();
        }

        protected override Window CreateWindow(IActivationState? activationState)
        {
            var window =  base.CreateWindow(activationState);
            window.Width = 1000;
            return window;

        }
    }
}
