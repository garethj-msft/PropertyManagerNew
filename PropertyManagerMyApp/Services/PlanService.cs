using Microsoft.Graph;
using Newtonsoft.Json.Linq;
using SuiteLevelWebApp.Utils;
using System;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Threading.Tasks;
using System.Collections.Generic;
using GraphModelsExtension;
using Newtonsoft.Json;

namespace SuiteLevelWebApp.Services
{
    public static class PlanService
    {
        private static async Task<Plan> CreatePlanAsync(Group group)
        {
            var accessToken = AuthenticationHelper.GetGraphAccessTokenAsync();

            var plansEndPoint = string.Format("{0}plans", AADAppSettings.GraphBetaResourceUrl);

            dynamic postPlanJSON = new JObject();
            postPlanJSON.title = group.DisplayName + " Plan";
            postPlanJSON.owner = group.Id;

            var requestMessage = new HttpRequestMessage(HttpMethod.Post, plansEndPoint);
            requestMessage.Content = new StringContent(postPlanJSON.ToString(), System.Text.Encoding.UTF8, "application/json");

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var responseMessage = await client.SendAsync(requestMessage);

                if (responseMessage.StatusCode != System.Net.HttpStatusCode.Created)
                    throw new Exception();

                var payload = await responseMessage.Content.ReadAsStringAsync();

                return new Plan
                {
                    id = JObject.Parse(payload)["id"].ToString()
                };
            }
        }

        public static async Task<Plan> GetPlanAsync(Group group)
        {
            var accessToken = AuthenticationHelper.GetGraphAccessTokenAsync();

            var plansEndPoint = string.Format("{0}groups/{1}/plans", AADAppSettings.GraphBetaResourceUrl, group.Id);

            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await accessToken);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    var responseMessage = await client.GetAsync(plansEndPoint);

                    if (responseMessage.StatusCode != System.Net.HttpStatusCode.OK)
                        throw new Exception();

                    var payload = await responseMessage.Content.ReadAsStringAsync();

                    var jobject = JObject.Parse(payload);

                    if (jobject["value"].Children().Count() > 0)
                    {
                        return new Plan
                        {
                            id = jobject["value"][0]["id"].ToString()
                        };
                    }
                    else
                    {
                        return await CreatePlanAsync(group);
                    }
                }
            }
            catch
            {
                return null;
            }
        }

        public static async Task<Bucket> GetBucketByNameAsync(Plan plan, string bucketName)
        {
            var accessToken = AuthenticationHelper.GetGraphAccessTokenAsync();

            var bucketsEndpoint = string.Format("{0}plans/{1}/buckets", AADAppSettings.GraphBetaResourceUrl, plan.id);

            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await accessToken);
                    client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                    var responseMessage = await client.GetAsync(bucketsEndpoint);

                    if (responseMessage.StatusCode != System.Net.HttpStatusCode.OK)
                        throw new Exception();

                    var payload = await responseMessage.Content.ReadAsStringAsync();

                    var jobject = JObject.Parse(payload);
                    var retVal = new Bucket
                    {
                        planId = plan.id,
                        name = bucketName
                    };
                    var matchingBuckets = jobject["value"].Children().Where(t => t["name"].ToString() == bucketName);
                    if (matchingBuckets.Count() > 0)
                    {
                        retVal.id = matchingBuckets.First()["id"].ToString();
                        return retVal;
                    }
                    else
                    {
                        return await CreateBucketAsync(retVal);
                    }
                }
            }
            catch
            {
                return null;
            }
        }


        public static async Task<Bucket> CreateBucketAsync(Bucket bucket)
        {
            var accessToken = AuthenticationHelper.GetGraphAccessTokenAsync();
            var bucketsEndPoint = string.Format("{0}buckets", AADAppSettings.GraphBetaResourceUrl);

            dynamic postBucketJSON = new JObject();
            postBucketJSON.name = bucket.name;
            postBucketJSON.planId = bucket.planId;

            var requestMessage = new HttpRequestMessage(HttpMethod.Post, bucketsEndPoint);
            requestMessage.Content = new StringContent(postBucketJSON.ToString(), System.Text.Encoding.UTF8, "application/json");

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var responseMessage = await client.SendAsync(requestMessage);

                if (responseMessage.StatusCode != System.Net.HttpStatusCode.Created)
                    throw new Exception();

                var payload = await responseMessage.Content.ReadAsStringAsync();

                var jobject = JObject.Parse(payload);

                return new Bucket
                {
                    id = jobject["id"].ToString(),
                    planId = jobject["planId"].ToString()
                };
            }
        }

        public static async Task<task> CreateTaskAsync(task task)
        {
            var accessToken = AuthenticationHelper.GetGraphAccessTokenAsync();
            var tasksEndPoint = string.Format("{0}tasks", AADAppSettings.GraphBetaResourceUrl);

            var requestMessage = new HttpRequestMessage(HttpMethod.Post, tasksEndPoint);
            requestMessage.Content = new StringContent(JsonConvert.SerializeObject(task,new JsonSerializerSettings { DefaultValueHandling = DefaultValueHandling.Ignore }), System.Text.Encoding.UTF8, "application/json");

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var responseMessage = await client.SendAsync(requestMessage);

                if (responseMessage.StatusCode != System.Net.HttpStatusCode.Created)
                    throw new Exception();

                var payload = await responseMessage.Content.ReadAsStringAsync();
                task = JsonConvert.DeserializeObject<task>(payload);
                task.etag = responseMessage.Headers.ETag;
                return task;
            }
        }

        public static async Task UpdateTaskDescriptionAsync(task task, string description)
        {
            if (task.etag == null) throw new ArgumentNullException("task.etag");

            var accessToken = AuthenticationHelper.GetGraphAccessTokenAsync();
            var tasksEndPoint = string.Format("{0}tasks/{1}/details", AADAppSettings.GraphBetaResourceUrl, task.id);

            var requestMessage = new HttpRequestMessage(new HttpMethod("PATCH"), tasksEndPoint);

            TaskDetails details = new TaskDetails { description = description, previewType = "description" };
            requestMessage.Content = new StringContent(JsonConvert.SerializeObject(details), System.Text.Encoding.UTF8, "application/json");

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.IfMatch.Add(new EntityTagHeaderValue(task.etag.Tag, task.etag.IsWeak));
                var responseMessage = await client.SendAsync(requestMessage);

                if (responseMessage.StatusCode != System.Net.HttpStatusCode.NoContent)
                    throw new Exception();
            }
        }

        public static async Task<task[]> GetTasksAsync(Plan plan)
        {
            var accessToken = AuthenticationHelper.GetGraphAccessTokenAsync();

            List<task> tasks = new List<task>();

            var plansEndPoint = string.Format("{0}plans/{1}/Tasks?filter=percentComplete+ne+100", AADAppSettings.GraphBetaResourceUrl, plan.id);

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", await accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                var responseMessage = await client.GetAsync(plansEndPoint);

                if (responseMessage.StatusCode != System.Net.HttpStatusCode.OK)
                    throw new Exception();

                var payload = await responseMessage.Content.ReadAsStringAsync();

                var jobject = JObject.Parse(payload);

                foreach (var item in jobject["value"].Children())
                {
                    tasks.Add(new task
                    {
                        title = item["title"].ToString(),
                        assignedTo = !string.IsNullOrEmpty(item["assignedTo"].ToString()) ? item["assignedTo"].ToString() : "",
                        assignedBy = !string.IsNullOrEmpty(item["assignedBy"].ToString()) ? item["assignedBy"].ToString() : "",
                        assignedDateTime = !string.IsNullOrEmpty(item["assignedDateTime"].ToString()) ? DateTimeOffset.Parse(item["assignedDateTime"].ToString()) : (DateTimeOffset?)null,
                        percentComplete = !string.IsNullOrEmpty(item["percentComplete"].ToString()) ? Convert.ToInt32(item["percentComplete"].ToString()) : 0
                    });
                }
            }

            return tasks.Where(i => i.percentComplete.HasValue && i.percentComplete.Value < 100).ToArray();
        }
    }
}