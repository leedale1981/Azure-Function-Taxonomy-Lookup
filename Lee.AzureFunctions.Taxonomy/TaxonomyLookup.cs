using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using Lee.AzureFunctions.Authentication;
using Lee.AzureFunctions.Constants;
using Lee.AzureFunctions.Helpers;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;

namespace Lee.AzureFunctions.Taxonomy
{
    public static class TaxonomyLookup
    {
        [FunctionName("TaxonomyLookup")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Function, 
            "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            string siteUrl = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "siteUrl", true) == 0)
                .Value;

            string siteCollectionUrl = siteUrl.ToLower().Substring(0, siteUrl.IndexOf("_layout"));

            log.Info($"C# HTTP trigger function looking up taxonomy value for site {siteCollectionUrl}.");

            string value = GetValueFromTermStore(siteCollectionUrl, log);

            return new HttpResponseMessage(HttpStatusCode.OK) { Content = new StringContent(value) };
        }

        private static string GetValueFromTermStore(string siteCollectionUrl, TraceWriter log)
        {
            string value = string.Empty;

            using (ClientContext context = new AuthenticationContext(siteCollectionUrl).GetAuthenticationContext())
            {
                TaxonomySession session = TaxonomySession.GetTaxonomySession(context);
                TermStore store = session.GetDefaultSiteCollectionTermStore();
                context.Load(store,
                            s => s.Name,
                            s => s.Groups.Include(
                                g => g.Name,
                                g => g.TermSets.Include(
                                    ts => ts.Name,
                                    ts => ts.CustomProperties,
                                    ts => ts.Terms.Include(
                                        t => t.Name,
                                        t => t.CustomProperties
                ))));

                context.ExecuteQueryRetry();
                log.Info($"Loaded taxonomy data");

                if (session != null && store != null)
                {
                    value = GetValueFromTermGroup(store, log, context);
                }
            }

            return value;
        }

        private static string GetValueFromTermGroup(TermStore store, TraceWriter log, ClientContext context)
        {
            string value = string.Empty;
            string[] termNames = EnvironmentConfigurationManager.GetSetting(AppSettings.TermGroupName).Split(';');
            string termGroupName = termNames[0];

            foreach (TermGroup group in store.Groups)
            {
                if (group.Name == termGroupName)
                {
                    log.Info($"Found {termGroupName} term group");
                    value = GetValueFromTermSet(group, termNames, log, context);
                }
            }

            return value;
        }

        private static string GetValueFromTermSet(TermGroup group, string[] termNames, TraceWriter log, ClientContext context)
        {
            string value = string.Empty;

            foreach (TermSet termSet in group.TermSets)
            {
                if (termSet.Name == termNames[1])
                {
                    log.Info($"Found {termNames[1]} term set");

                    foreach (Term term in termSet.Terms)
                    {
                        if (term.Name == termNames[2])
                        {
                            context.Load(term.Terms, t => t.Include(
                            st => st.Name,
                            st => st.LocalCustomProperties));
                            context.ExecuteQueryRetry();

                            log.Info($"Found {term.Name} term");
                            value = term.LocalCustomProperties["Value"];
                            log.Info($" {termNames[2]} is {value}");
                        }
                    }
                    
                }
            }

            return value;
        }
    }
}
