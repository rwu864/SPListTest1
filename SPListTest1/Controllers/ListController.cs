using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;
using SPClient = Microsoft.SharePoint.Client;
using Microsoft.SharePoint;
using Microsoft.SharePoint.Utilities;


namespace SPListTest1.Controllers
{
    public class ListController : Controller
    {
        string SiteUrl = "http://qtcserver/raymond/";
        string ListName = "Request_Test";

        public ActionResult Items()
        {
            ClientContext clientContext = new ClientContext(SiteUrl);

            // The object model does not automatically load data so we need to
            // use queries to retrieve data

            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            clientContext.Load(spList);
            clientContext.ExecuteQuery();

            //Internal Names                ->  List Names 
            //NewColumn1                    ->  Request ID
            //Request_x0020_Details         ->  Request Details
            //Author                        ->  Request By
            //Request_x0020_Status          ->  Request Status
            //Request_x0020_Due_x0020_Date  -> Request Due Date

            CamlQuery camlQuery = new CamlQuery();
            ListItemCollection spListItems = spList.GetItems(camlQuery);
            clientContext.Load(spListItems, items => items.IncludeWithDefaultProperties(
                item => item["NewColumn1"],
                item => item["Request_x0020_Details"],
                item => item["Author"],
                item => item["Request_x0020_Status"],
                item => item["Request_x0020_Due_x0020_Date"]));
            clientContext.ExecuteQuery();

            ViewBag.List = spList;
            ViewBag.ListItems = spListItems;

            

            
            return View();
        }

        [HttpGet]
        public ActionResult NewItem()
        {
            return View();
        }

        [HttpPost]
        public RedirectResult NewItem(string details, DateTime due_date)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);

            // For testing purpose, we'll use a local account jdoe for the new item;
            // otherwise the Requested By (i.e. Author) field will be populated with
            // the user who runs the ASP.NET web applicationn.

            var user = clientContext.Web.EnsureUser("rwu8");
            clientContext.Load(user);
            clientContext.ExecuteQuery();
            FieldUserValue userValue = new FieldUserValue();
            userValue.LookupId = user.Id;

            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            var info = new ListItemCreationInformation();
            var item = spList.AddItem(info);
            item["Request_x0020_Details"] = details;
            item["Author"] = userValue;
            item["Request_x0020_Due_x0020_Date"] = due_date;
            item.Update();
            clientContext.ExecuteQuery();

            return Redirect("Items");
        }

        [HttpGet]
        public ActionResult DeleteItem(int ID)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);
            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            ListItem spListItem = spList.GetItemById(ID);

            spListItem.DeleteObject();
            clientContext.ExecuteQuery();

            return Redirect("Items");
        }

        [HttpGet]
        public ActionResult EditItem(int ID)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);
            List spList = clientContext.Web.Lists.GetByTitle(ListName);

            clientContext.Load(spList);
            ListItem spListItem = spList.GetItemById(ID);
            clientContext.Load(spListItem);
            clientContext.ExecuteQuery();
            
            string Request_ID = (String)spListItem["NewColumn1"];
            string Request_Details = (String)spListItem["Request_x0020_Details"];
            string Request_Status = (String)spListItem["Request_x0020_Status"];
            FieldUserValue Author = (FieldUserValue)spListItem["Author"];
            string Request_By = Author.LookupValue;
            string Request_Due_Date = ((DateTime)spListItem["Request_x0020_Due_x0020_Date"]).ToShortDateString();
            
            ViewBag.Request_ID = Request_ID;
            ViewBag.Request_Details = Request_Details;
            ViewBag.Request_Status = Request_Status;
            ViewBag.Request_By = Request_By;
            ViewBag.ID = ID;
            ViewBag.Request_Due_Date = Request_Due_Date;

            //retrieving drop down menu items 
            var choices_list = spListItem["Request_x0020_Status"].ToString();
            ViewBag.Choices_List = choices_list;

            return View();
        }

        [HttpPost]
        public RedirectResult EditItem(String details, String status, int ID, DateTime date)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);
            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            ListItem spListItem = spList.GetItemById(ID);

            spListItem["Request_x0020_Details"] = details;
            spListItem["Request_x0020_Status"] = status;
            spListItem["Request_x0020_Due_x0020_Date"] = date;

            spListItem.Update();
            clientContext.ExecuteQuery();

            return Redirect("Items");
        }
    }
}