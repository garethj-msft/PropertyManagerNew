using GraphModelsExtension;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SuiteLevelWebApp.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Web;

namespace SuiteLevelWebApp.Services
{
    public static class SharePointService
    {
        public static async Task<Site> GetSiteByPathAsync(string sitePath)
        {
            // https://graph.microsoft.com/beta/sharePoint:{SITEPATH}
            var accessToken = await AuthenticationHelper.GetGraphAccessTokenAsync();
            var pathQueryEndPoint = $"{AADAppSettings.GraphBetaResourceUrl}sharePoint:{sitePath}";

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var responseMessage = await client.GetAsync(pathQueryEndPoint);

                if (responseMessage.StatusCode != System.Net.HttpStatusCode.OK)
                    throw new Exception();

                var payload = await responseMessage.Content.ReadAsStringAsync();
                Site site = JsonConvert.DeserializeObject<Site>(payload);
                return site;
            }
        }

        internal static async Task<T[]> GetListItemsFromSiteList<T>(Site site, object listName)
        {
            // GET https://graph.microsoft.com/beta/sharepoint/sites/{siteId}/lists?filter=name%20eq%20'{listname}'&expand=items(expand=columnSet)&select=items HTTP/1.1
            var accessToken = await AuthenticationHelper.GetGraphAccessTokenAsync();
            var pathQueryEndPoint = $"{AADAppSettings.GraphBetaResourceUrl}sharePoint/sites/{site.id}/lists?filter=name eq '{listName}'&expand=items(expand=columnSet)&select=items";

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var responseMessage = await client.GetAsync(pathQueryEndPoint);

                if (responseMessage.StatusCode != System.Net.HttpStatusCode.OK)
                    throw new Exception();

                var payload = await responseMessage.Content.ReadAsStringAsync();
                List[] lists = JsonConvert.DeserializeObject<ValueWrapper<List[]>>(payload).value;
                return lists.FirstOrDefault()?.items.Select(i => i.columnSet.ToObject<T>()).ToArray();
            }
        }
    }
}