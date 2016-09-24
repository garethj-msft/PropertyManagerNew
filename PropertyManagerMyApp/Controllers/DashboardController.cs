using SuiteLevelWebApp.Models;
using SuiteLevelWebApp.Services;
using SuiteLevelWebApp.Utils;
using System;
using System.Configuration;
using System.Web;
using System.Threading.Tasks;
using System.Web.Mvc;
using Microsoft.Graph;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Web.Http.Results;
using Newtonsoft.Json;
using System.Collections.Generic;
using GraphModelsExtension;

namespace SuiteLevelWebApp.Controllers
{
    [Authorize, HandleAdalException]
    public class DashboardController : Controller
    {
        public async Task<ActionResult> Index()
        {
            var sharePointToken = await AuthenticationHelper.GetAccessTokenAsync(AppSettings.DemoSiteServiceResourceId);
            Dashboard dashboardModel = new Dashboard(sharePointToken);
            var model = await dashboardModel.GetDashboardPropertiesViewModelAsync();
            return View(model);
        }

        public async Task<ActionResult> Property(int id)
        {
            var sharePointToken = await AuthenticationHelper.GetAccessTokenAsync(AppSettings.DemoSiteServiceResourceId);
            Dashboard dashboardModel = new Dashboard(sharePointToken);
            var model = await dashboardModel.GetDashboardPropertiesViewModelAsync(id);

            return View(model);
        }

        public async Task<ActionResult> AddInspector()
        {
            var sharePointToken = await AuthenticationHelper.GetAccessTokenAsync(AppSettings.DemoSiteServiceResourceId);
            var graphService = await AuthenticationHelper.GetGraphServiceAsync(AADAppSettings.GraphResourceUrl);
            Dashboard dashboardModel = new Dashboard(sharePointToken);
            var model = await dashboardModel.GetDashboardAddInspectorViewModelAsync(graphService);
            return View(model);
        }

        public async Task<ActionResult> AddInspectorExecute(AddInspectorViewModel viewModel)
        {
            // Acquire all the fixed resources we'll need.
            var graphService = await AuthenticationHelper.GetGraphServiceAsync(AADAppSettings.GraphResourceUrl);
            User candidate = await graphService.Users[viewModel.SelectedCandidate].Request().GetAsync();
            Group inspectors = await graphService.GetGroupByDisplayNameAsync("Inspectors");
            Group gettingStarted = await graphService.GetGroupByDisplayNameAsync("GettingStarted");
            Plan gettingStartedPlan = await PlanService.GetPlanAsync(gettingStarted);
            Bucket bucket = await PlanService.GetBucketByNameAsync(gettingStartedPlan, "Unstarted");

            // Add the user to the GettingStarted group where newcomers share onboarding tasks
            try
            {
                await graphService.Groups[gettingStarted.Id].Members.References.Request().AddAsync(candidate);
            }
            catch
            {
                // Assume failure means it's already in the group - as it often is during demos!
            }
 

            // Iterate over the 'NewHireTasks' sharepoint list with the new preview API.
            var site = await SharePointService.GetSiteByPathAsync("/sites/SuiteLevelAppDemo");

            NewHireTaskListitem[] listitems = await SharePointService.GetListItemsFromSiteList<NewHireTaskListitem>(site, "NewHireTasks");

            foreach (NewHireTaskListitem listItem in listitems)
            {
                try
                {
                    // Make a Planner task in the 'GettingStarted' group for each task in the NewHireTasks list.
                    task created = await PlanService.CreateTaskAsync(new task
                    {
                        title = listItem.Title,
                        assignedTo = candidate.Id,
                        percentComplete = 0,
                        planId = bucket.planId,
                        bucketId = bucket.id,
                    });
                    await PlanService.UpdateTaskDescriptionAsync(created, listItem.Description);
                }
                catch
                {
                    // Move on to the next if anything goes wrong.
                    continue;
                }
            }

            // Make sure new user has all the licenses they need.
            await GraphServiceExtension.AssignLicenseAsync(graphService, candidate);

            // Add the user to the Inspector's group
            await graphService.Groups[inspectors.Id].Members.References.Request().AddAsync(candidate);

            return View(viewModel);
        }


        public async Task<ActionResult> InspectionDetails(int id)
        {
            var sharePointToken = await AuthenticationHelper.GetAccessTokenAsync(AppSettings.DemoSiteServiceResourceId);
            var graphService = await AuthenticationHelper.GetGraphServiceAsync(AADAppSettings.GraphResourceUrl);
            Dashboard dashboardModel = new Dashboard(sharePointToken);
            var model = await dashboardModel.GetDashboardInspectionDetailsViewModelAsync(graphService, id, User.Identity.Name);
            if (model == null) return HttpNotFound();
            var accessToken = await AuthenticationHelper.GetGraphAccessTokenAsync();
            TempData["accesstoken"] = accessToken;
            await dashboardModel.CheckSubscriptionAsync(graphService);
            return View(model);
        }

        
        [HttpGet]
        public async Task<FileResult> GetPropertyGroupExcelChart(string groupId, string fileId, string title)
        {
            var imageBinary = await ExcelService.GetPropertyExcelWorkbookChartImageAsync(groupId, fileId, title);
            return new FileContentResult(imageBinary, "image/bmp"); ;
        }

        [HttpPost]
        public async Task<ActionResult> ScheduleRepair(ScheduleRepairModel model)
        {
            model.TimeSlotsSelectedValue = DateTime.ParseExact(string.Format("{0} {1}:00", model.TimeSlotsSelectedDateValue, model.TimeSlotsSelectedHoursValue),
                "yyyy-MM-dd HH:mm", null);
            var sharePointToken = AuthenticationHelper.GetAccessTokenAsync(AppSettings.DemoSiteServiceResourceId);
            var graphService = AuthenticationHelper.GetGraphServiceAsync(AADAppSettings.GraphResourceUrl);

            var tasksService = new TasksService(await sharePointToken);
            var dashboardModel = new Dashboard(await sharePointToken);

            await dashboardModel.UpdateRepairScheduleInfoToIncidentAsync(model)
                .ContinueWith(async task => await tasksService.CompleteRepairAssignmentTaskAsync(model.IncidentId));

            await dashboardModel.ScheduleRepairAsync(await graphService, model);
            await dashboardModel.CreateGroupRepairEventAsync(await graphService, model);
            await dashboardModel.CreateO365TaskAsync(await graphService, model);

            return RedirectToAction("Index");
        }

        [HttpPost]
        public async Task<ActionResult> AuditRepair(AuditRepairModel model)
        {
            var sharePointToken = await AuthenticationHelper.GetAccessTokenAsync(AppSettings.DemoSiteServiceResourceId);
            Dashboard dashboardModel = new Dashboard(sharePointToken);

            var tasksService = new TasksService(sharePointToken);
            await tasksService.CompleteRepairApprovalTaskAsync(model);

            if (model.Result == ApprovalResult.Approved)
                await dashboardModel.ApproveRepairAsync(model.IncidentId);

            return RedirectToAction("Index");
        }

        [HttpPost]
        public async Task<ActionResult> AnnotateImages(int incidentId)
        {
            var sharePointToken = AuthenticationHelper.GetAccessTokenAsync(AppSettings.DemoSiteServiceResourceId);
            var graphService = AuthenticationHelper.GetGraphServiceAsync(AADAppSettings.GraphResourceUrl);
            var dashboardService = new Dashboard(await sharePointToken);

            var pageUrl = await dashboardService.AnnotateImagesAsync(await graphService, Server.MapPath("/"), incidentId);
            return Redirect(pageUrl);
        }

        [HttpPost]
        public async Task<ActionResult> UploadFile(UploadFileModel model)
        {
            if (model.File != null)
            {
                var token = AuthenticationHelper.GetAccessTokenAsync(AppSettings.DemoSiteServiceResourceId);
                var graphService = AuthenticationHelper.GetGraphServiceAsync(AADAppSettings.GraphResourceUrl);

                var dashboardService = new Dashboard(await token);
                await dashboardService.UploadFileAsync(await graphService, model);
            }

            return RedirectToAction("InspectionDetails", new { id = model.IncidentId });
        }

    }
}