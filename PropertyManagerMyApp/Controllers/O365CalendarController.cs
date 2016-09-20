﻿using SuiteLevelWebApp.Models;
using SuiteLevelWebApp.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft.Graph;
using System.Net.Http;
using System.Net.Http.Headers;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace SuiteLevelWebApp.Controllers
{
    [Authorize, HandleAdalException]
    public class O365CalendarController : Controller
    {
        //
        //https://msdn.microsoft.com/en-us/office/office365/APi/calendar-rest-operations#Findmeetingtimespreview
        //
        public async Task<JsonResult> GetAvailableTimeSlots(string localDate, string userEmail)
        {
            List<TimeSlot> timeSlots = new List<TimeSlot>();
            var accessToken = await AuthenticationHelper.GetGraphAccessTokenAsync();

            var restURL = "https://graph.microsoft.com/beta/me/findMeetingTimes";
            object[] attendeeArray = { new { type = "Required", emailAddress = new { address = userEmail } } };
            var timeZone = TimeZone.CurrentTimeZone.StandardName;//"Pacific Standard Time";
            object[] TimeConstraintArray = { new { start = new { date = localDate, time = "9:00:00", timeZone =  timeZone},
                                                   end = new { date = localDate, time = "17:00:00", timeZone = timeZone } } };
            var TimeConstraint = new { timeslots = TimeConstraintArray/*, ActivityDomain = "Personal" */};

            var requstBody = new { attendees = attendeeArray, timeConstraint = TimeConstraint, meetingDuration = "PT1H" , MaxCandidates = "10"};
            var requestMessage = new HttpRequestMessage(HttpMethod.Post, restURL);
            string contentString = JsonConvert.SerializeObject(requstBody, new JsonSerializerSettings { NullValueHandling = NullValueHandling.Ignore });
            requestMessage.Content = new StringContent(contentString, System.Text.Encoding.UTF8, "application/json");

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);
                client.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                client.DefaultRequestHeaders.Add("Prefer", "outlook.timezone=\"" + timeZone + "\"");
                HttpResponseMessage responseMessage = await client.SendAsync(requestMessage);

                if (responseMessage.IsSuccessStatusCode)
                {
                    string stringResult = await responseMessage.Content.ReadAsStringAsync();
                    JObject ret = JObject.Parse(stringResult);
                    JArray array = (JArray)ret["meetingTimeSlots"];
                    foreach (var item in array)
                    {
                        if (item["meetingTimeSlot"] != null)
                        {
                            var start = item["meetingTimeSlot"]["start"];
                            var end = item["meetingTimeSlot"]["end"];

                            AddTimeSlot(timeSlots, start["time"].ToString(), end["time"].ToString());
                        }
                    }
                }
            }
            return Json(timeSlots, JsonRequestBehavior.AllowGet);
        }
        private void AddTimeSlot(List<TimeSlot> timeSlots, string start, string end)
        {
            var availableStart = TimeSpan.Parse(start);
            var availableEnd = TimeSpan.Parse(end);

            TimeSlot timeSlot = new TimeSlot
            {
                Start = availableStart.Hours > 9 ? availableStart.Hours.ToString() : string.Format("0{0}", availableStart.Hours),
                Value = string.Format("{0} - {1}", availableStart.ToString(@"hh\:mm"), availableEnd.ToString(@"hh\:mm"))
            };
            timeSlots.Add(timeSlot);
        }
    }
}