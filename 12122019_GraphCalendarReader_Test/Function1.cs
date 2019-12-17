using System;
using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;
using Newtonsoft.Json;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using Microsoft.Graph.Auth;
using System.Text;
using System.Collections.Generic;

namespace _12122019_GraphCalendarReader_Test
{
    public static class Function1
    {
        [FunctionName("Function1")]
        public static async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequest req, ILogger log)
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


            StringBuilder s = new StringBuilder();

            foreach (var user in users)
            {
                
                if (user.DisplayName != "Bianca Pisani"
                    &&user.DisplayName != "Cameron White"
                    &&user.DisplayName != "Delia Dennis"
                    &&user.DisplayName != "Gerhart Moller"
                    &&user.DisplayName != "Provisioning User"
                    &&user.DisplayName != "Raul Razo")
                {
                    var events = await graphClient.Users[user.Id].Events.Request()
                        .Expand("Calendar")
                        //.Filter("categories/any(a:a eq '" + "ziekte" + "')")
                        .GetAsync();
                    
                    if (events.CurrentPage.Count>0)
                    {
                        s.AppendLine(user.DisplayName);
                        for (int i = 0; i < events.CurrentPage.Count; i++)
                        {
                            var thisEvent = events[i];
                            var categories = thisEvent.Categories;
                            var startDatum = Convert.ToDateTime(thisEvent.Start.DateTime).ToString("dddd/MM/yyyy");
                            var eindDatum = Convert.ToDateTime(thisEvent.End.DateTime).ToString("dddd/MM/yyyy");
                            var startTijd = Convert.ToDateTime(thisEvent.Start.DateTime).ToString("H:mm");
                            var eindTijd = Convert.ToDateTime(thisEvent.End.DateTime).ToString("H:mm");
                            string alleCategorieen = "";
                            foreach (var item in categories)
                            {
                                alleCategorieen +=" " + item.ToString();
                            }                

                            s.AppendLine($"             {startDatum.PadLeft(17)} {thisEvent.Calendar.Name.PadLeft(20)}  {alleCategorieen.PadLeft(20)}   {thisEvent.Subject.PadLeft(50)}       {startTijd}          {eindTijd}");                            
                        }
                    }
                }
            }            

    
                           string result = "";
                return result != null
        ? (ActionResult)new OkObjectResult($"{s}")
        : new BadRequestObjectResult("Please pass a name on the query string or in the request body");
            }


    }
    }



