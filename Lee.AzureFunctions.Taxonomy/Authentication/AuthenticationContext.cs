using Lee.AzureFunctions.Constants;
using Lee.AzureFunctions.Helpers;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core;
using System.IO;

namespace Lee.AzureFunctions.Authentication
{
    public class AuthenticationContext
    {
        private readonly string homePath =
            EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.HomeRootPath);
        private readonly string certName = 
            EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.CertName);
        private string certPath = string.Empty;
        private string siteUrl = string.Empty;

        public AuthenticationContext()
        {
            this.siteUrl =
                EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.TenantAdminUrl);
            this.certPath =
                Path.Combine(this.homePath, this.homePath.ToLower().IndexOf("c:") > -1 ? "Cert" : "site\\wwwroot\\Cert");
        }

        public AuthenticationContext(string siteUrl) : base()
        {
            this.siteUrl = siteUrl;
            this.certPath =
                Path.Combine(this.homePath, this.homePath.ToLower().IndexOf("c:") > -1 ? "Cert" : "site\\wwwroot\\Cert");
        }

        public ClientContext GetAuthenticationContext()
        {
            var authType = this.GetAuthType();
            ClientContext authContext = null;

            switch (authType)
            {
                case AuthenticationType.AppOnly:
                    authContext = this.GetAppOnlyAuthContext(siteUrl);
                    break;

                case AuthenticationType.SPAuth:
                    authContext = this.GetSPAuthContext(siteUrl);
                    break;
            }

            return authContext;
        }

        private ClientContext GetSPAuthContext(string siteUrl)
        {
            string clientId =
                EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.AuthId);
            string clientSecret =
                EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.TenantDomain);
            string certPassword =
                EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.Password);

            return new AuthenticationManager().GetAppOnlyAuthenticatedContext(
                siteUrl, clientId, clientSecret);
        }

        private ClientContext GetAppOnlyAuthContext(string siteUrl)
        {
            string clientId =
                EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.AuthId);
            string tenant =
                EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.TenantDomain);
            string certPassword =
                EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.Password);

            return new AuthenticationManager().GetAzureADAppOnlyAuthenticatedContext(
                siteUrl, clientId, tenant, Path.Combine(this.certPath, this.certName), certPassword);
        }

        private AuthenticationType GetAuthType()
        {
            string authTypeSetting =
                EnvironmentConfigurationManager.GetEnvironmentVariable(AppSettings.AuthType);

            switch (authTypeSetting)
            {
                case "AppOnly":
                    return AuthenticationType.AppOnly;

                case "SPAuth":
                    return AuthenticationType.SPAuth;

                default:
                    return AuthenticationType.SPAuth;
            }
        }
    }
}

