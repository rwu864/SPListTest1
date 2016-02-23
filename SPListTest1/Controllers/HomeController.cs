using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;
using SPClient = Microsoft.SharePoint.Client;

namespace SPListTest1.Controllers
{
    public class HomeController : Controller
    {
        string SiteUrl = "http://win08vm/cysun";
        string ListName = "Test List 1";

        public ActionResult Index()
        {
            ClientContext clientContext = new ClientContext(SiteUrl);

            // The object model does not automatically load data so we need to
            // use queries to retrieve data

            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            clientContext.Load(spList);
            clientContext.ExecuteQuery();

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection spListItems = spList.GetItems(camlQuery);
            clientContext.Load(spListItems, items => items.IncludeWithDefaultProperties(
                item => item["Request_x0020_ID"],
                item => item["Request_x0020_Details"],
                item => item["Author"],
                item => item["Request_x0020_Status"]));
            clientContext.ExecuteQuery();

            ViewBag.List = spList;
            ViewBag.ListItems = spListItems;
            return View();
        }
    }
}