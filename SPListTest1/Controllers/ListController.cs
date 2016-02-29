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

        //Internal Names                ->  List Names 
        //NewColumn1                    ->  Request ID
        //Request_x0020_Details         ->  Request Details
        //Author                        ->  Request By
        //Request_x0020_Status          ->  Request Status
        //Request_x0020_Due_x0020_Date  -> Request Due Date

        //these are strings that represent that map the SP2010 list name to the internal names 
    
        string sp_Request_ID = "NewColumn1";
        string sp_Request_Details = "Request_x0020_Details";
        string sp_Request_By = "Author";
        string sp_Request_Status = "Request_x0020_Status";
        string sp_Request_Due_Date = "Request_x0020_Due_x0020_Date";

        public ActionResult Items()
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
                item => item[sp_Request_ID],
                item => item[sp_Request_Details],
                item => item[sp_Request_By],
                item => item[sp_Request_Status],
                item => item[sp_Request_Due_Date]));
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
            item[sp_Request_Details] = details;
            item[sp_Request_By] = userValue;
            item[sp_Request_Due_Date] = due_date;
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
            
            string Request_ID = (String)spListItem[sp_Request_ID];
            string Request_Details = (String)spListItem[sp_Request_Details];
            string Request_Status = (String)spListItem[sp_Request_Status];
            FieldUserValue Author = (FieldUserValue)spListItem[sp_Request_By];
            string Request_By = Author.LookupValue;
            string Request_Due_Date = ((DateTime)spListItem[sp_Request_Due_Date]).ToShortDateString();
            
            ViewBag.Request_ID = Request_ID;
            ViewBag.Request_Details = Request_Details;
            ViewBag.Request_Status = Request_Status;
            ViewBag.Request_By = Request_By;
            ViewBag.ID = ID;
            ViewBag.Request_Due_Date = Request_Due_Date;

            return View();
        }

        [HttpPost]
        public RedirectResult EditItem(String details, String status, int ID, DateTime date)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);
            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            ListItem spListItem = spList.GetItemById(ID);

            spListItem[sp_Request_Details] = details;
            spListItem[sp_Request_Status] = status;
            spListItem[sp_Request_Due_Date] = date;

            spListItem.Update();
            clientContext.ExecuteQuery();

            return Redirect("Items");
        }
    }
}