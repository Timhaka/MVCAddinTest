using Microsoft.SharePoint.Client;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.Script.Serialization;

namespace MVCAddinTestWeb.Controllers
{
    public class HomeController : Controller
    {
        [SharePointContextFilter]
        public ActionResult Index(string SPHostUrl)
        {
            User spUser = null;

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (ClientContext ctx = spContext.CreateUserClientContextForSPHost())
            {
                if (ctx != null)
                {
                    spUser = ctx.Web.CurrentUser;

                    ctx.Load(spUser, user => user.Title);

                    ctx.ExecuteQuery();

                    ViewBag.UserName = spUser.Title;
                }

                ListCollection list = ctx.Web.Lists;
                ctx.Load(list);
                ctx.ExecuteQuery();


                ViewBag.spHost = SPHostUrl;
                ViewBag.Alllists = list;

            }

            



            return View();
        }

        public ActionResult ShowListItems(string listid, string SPHostUrl)
        {
        

            var spContext = SharePointContextProvider.Current.GetSharePointContext(HttpContext);

            using (ClientContext ctx = spContext.CreateUserClientContextForSPHost())
            {
                List listName = ctx.Web.Lists.GetById(listid.ToGuid());

                ctx.Load(listName);
                ctx.ExecuteQuery();

                ListItemCollection list = listName.GetItems(CamlQuery.CreateAllItemsQuery());
                List<String> listitemNames = new List<String>();
                ctx.Load(list);
                ctx.ExecuteQuery();


                    


                foreach (ListItem item in list)
                {
                    ctx.Load(item, include => include["Title"]);
                   // listitemNames.Add(item["Title"].ToString());
                }
                ctx.ExecuteQuery();

                foreach (ListItem item in list)
                {
                    listitemNames.Add(item["Title"].ToString());
                }

                string json = JsonConvert.SerializeObject(listitemNames);

                ViewBag.listItems = list;
                ViewBag.Listname = listName.Title;
                ViewBag.JsonTest = json;



                return View(list);


            }   
        }

            public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}
