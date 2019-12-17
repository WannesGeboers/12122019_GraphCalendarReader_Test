using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using _12122019_GraphCalendarReader_Test.Models;
using System.Text;
using System.Collections.Generic;
using Microsoft.Identity.Client;
using Microsoft.Graph.Auth;
using Microsoft.Graph;

namespace _12122019_GraphCalendarReader_Test
{
    public static class SaveCategoriesInDB
    {
        [FunctionName("SaveCategoriesInDB")]
        public static async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req,
            ILogger log)
        {
            var cronosADConfidentialClientApplication = ConfidentialClientApplicationBuilder
                 //.Create(config["CronosAD:ClientId"])
                 //.WithTenantId(config["CronosAD:TenantId"])
                 //.WithClientSecret(config["CronosAD:ClientSecret"])
                 .Create("61b395a0-ba37-4197-a91f-e1b1d2993b75")
                 .WithTenantId("a3a8360d-012a-451c-ae34-22db1f1754c8")
                 .WithClientSecret("w6AL/G551?ImZb.Nhv]Mw2vumtZ?diAR")
                 .Build();
            var cronosADAuthenticationProvider = new ClientCredentialProvider(cronosADConfidentialClientApplication);
            var graphClient = new GraphServiceClient(cronosADAuthenticationProvider);
            var users = await graphClient.Users.Request().GetAsync();


            List<EventItem> ListEventItems = new List<EventItem>();

            foreach (var user in users)
            {

                if (user.DisplayName != "Bianca Pisani"
                    && user.DisplayName != "Cameron White"
                    && user.DisplayName != "Delia Dennis"
                    && user.DisplayName != "Gerhart Moller"
                    && user.DisplayName != "Provisioning User"
                    && user.DisplayName != "Raul Razo")
                {
                    var events = await graphClient.Users[user.Id].Events.Request()
                        .Expand("Calendar")
                        //.Filter("categories/any(a:a eq '" + "ziekte" + "')")
                        .GetAsync();

                    if (events.CurrentPage.Count > 0)
                    {
                        for (int i = 0; i < events.CurrentPage.Count; i++)
                        {
                            var thisEvent = events[i];
                            var categories = thisEvent.Categories;

                            foreach (var item in categories)
                            {
                                if (item == "Ziekte"||item=="Vakantie"||item=="xxx")
                                {
                                    EventItem newEvent = new EventItem();
                                    newEvent.UserName = user.DisplayName;
                                    newEvent.StartDate = Convert.ToDateTime(thisEvent.Start.DateTime);
                                    newEvent.EndDate = Convert.ToDateTime(thisEvent.End.DateTime);
                                    newEvent.Subject = thisEvent.Subject;
                                    newEvent.Categories = thisEvent.Categories;



                                    ListEventItems.Add(newEvent);
                                }
                            }
                        }
                    }
                }
            }



            return ListEventItems != null
            ? (ActionResult)new OkObjectResult($"{Display(ListEventItems)}")
            : new BadRequestObjectResult("Good job, u broke it!");
        }

        private static string Display(List<EventItem> listEventItems)
        {

            StringBuilder result = new StringBuilder();
            result.AppendLine($"{"USERNAME".PadLeft(20)} {"DATUM".PadLeft(17)} {"START".PadLeft(17)} {"EINDE".PadLeft(17)} {"REDEN".PadLeft(30)} {"CATEGORIEEN".PadLeft(20)}");
            foreach (var item in listEventItems)
            {

                var startDatum = Convert.ToDateTime(item.StartDate).ToString("dd/MM/yyyy");
                var eindDatum = Convert.ToDateTime(item.EndDate).ToString("dd/MM/yyyy");
                var startTijd = Convert.ToDateTime(item.StartDate).ToString("H:mm");
                var eindTijd = Convert.ToDateTime(item.EndDate).ToString("H:mm");
                string alleCategorieen = "";
                foreach (var cat in item.Categories)
                {
                    alleCategorieen += " " + cat.ToString();
                }

                result.AppendLine($" {item.UserName.PadLeft(20)} {startDatum.PadLeft(17)} {startTijd.PadLeft(17)} {eindTijd.PadLeft(17)} {item.Subject.PadLeft(30)} {alleCategorieen.PadLeft(20)}        ");

            }

            return result.ToString();
        }
    }
}
