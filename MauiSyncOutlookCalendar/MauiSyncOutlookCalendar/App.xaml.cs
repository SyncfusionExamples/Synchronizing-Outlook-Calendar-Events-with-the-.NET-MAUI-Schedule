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
            clientID = "";

            //// You need to replace your tenant ID
            tenantID = "";

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
