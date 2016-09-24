using Microsoft.Graph;
using SuiteLevelWebApp.Utils;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Linq.Expressions;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using utils = SuiteLevelWebApp.Utils;
using System.Net;
using System.IO;

namespace SuiteLevelWebApp
{
    public static class GraphServiceExtension
    {
        //Extension methods for the Microsoft.Graph.GraphService returned
        //by the GetGraphServiceAsync() method in the AuthenticationHelper class.

        //Controllers use these extension methods to interact with the Microsoft.Graph.GraphService

        public static async Task<User[]> GetAllUsersAsync(this GraphServiceClient service, IEnumerable<string> displayNames)
        {
            var users = await service.Users.Request().GetAllAsync();
            var retUsers = users.Where(x => displayNames.Contains(x.DisplayName)).ToArray();

            return retUsers;
        }

        public static async Task<Group> GetGroupByDisplayNameAsync(this GraphServiceClient service, string displayName)
        {
            var groups = (await service.Groups.Request().Filter(string.Format("displayName eq '{0}'", displayName)).Top(1).GetAsync()).CurrentPage;
            return groups.Count > 0 ? groups[0] : null;
        }

        public static async Task<Group> AddGroupAsync(this GraphServiceClient service, string Name, string DisplayName, string Description)
        {
            Group newGroup = new Group
            {
                DisplayName = DisplayName,
                Description = Description,
                MailNickname = Name,
                MailEnabled = false,
                SecurityEnabled = true
            };
            try
            {
                newGroup = await service.Groups.Request().AddAsync(newGroup);
                return newGroup;
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        public static async Task<User[]> GetGroupMembersAsync(this GraphServiceClient service, string groupName)
        {
            var group = await GetGroupByDisplayNameAsync(service, groupName);

            if (group == null) return new User[0];

            var groupFetcher = service.Groups[group.Id];
            return await GetGroupMembersAsync(service, groupFetcher);
        }

        public static async Task<User[]> GetGroupMembersAsync(this GraphServiceClient service, IGroupRequestBuilder groupFetcher)
        {
            List<User> users = new List<User>();
            var collection = await groupFetcher.Members.Request().GetAllAsync();
            foreach (var item in collection)
            {
                var findUser = await service.Users[item.Id].Request().Select("id,displayName,department,officeLocation,mail,mobilePhone,businessPhones,jobTitle").GetAsync();
                if (findUser != null)
                    users.Add(findUser);
            }

            return users.ToArray();
        }

        public static async Task AssignLicenseAsync(this GraphServiceClient service, User user)
        {
            var subscribedSkus = await service.SubscribedSkus.Request().GetAllAsync();
          
            foreach (SubscribedSku sku in subscribedSkus)
            {
                if ((sku.CapabilityStatus == "Enabled") &&
                     (sku.PrepaidUnits.Enabled.Value > sku.ConsumedUnits))
                {
                    user = await service.Users[user.Id].AssignLicense(new[] { new AssignedLicense { SkuId = sku.SkuId.Value } }, new Guid[] { }).Request().PostAsync();
                }
            }
        }

        public static async Task<Subscription> CreateSubscriptionAsync(this GraphServiceClient service)
        {
            Subscription subscription = new Subscription()
            {
                ChangeType = "created,updated",
                ClientState = Guid.NewGuid().ToString(),
                //https://graph.microsoft.io/en-us/docs/api-reference/v1.0/resources/subscription
                ExpirationDateTime = DateTime.UtcNow + new TimeSpan(0, 2, 0, 0),//TimeSpan.FromMinutes(4230),//
                NotificationUrl = ConfigurationManager.AppSettings["NotificationUrl"],
                Resource = "me/mailFolders('Inbox')/messages"
            };
            int retrycount = 5;
            while (retrycount-- > 0)
            {
                try
                {
                    subscription = await service.Subscriptions.Request().AddAsync(subscription);
                    break;
                }
                catch
                {

                }
            }
            return subscription.Id != null ? subscription : null;
        }
        public static async Task DeleteSubscriptionAsync(this GraphServiceClient service, string subscriptionId)
        {
            try
            {
                await service.Subscriptions[subscriptionId].Request().DeleteAsync();
                return;
            }
            catch (Exception)
            {
                //throw ex;
            }
        }
    }
}