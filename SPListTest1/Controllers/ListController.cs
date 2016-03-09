using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using Microsoft.SharePoint.Client;
using SPClient = Microsoft.SharePoint.Client;

namespace SPListTest1.Controllers
{
    public class ListController : Controller
    {
        string SiteUrl = "http://qtcserver/raymond";
        string ListName = "Request_Test";

        // Internal names of the fields
        string FieldId = "NewColumn1";
        string FieldDetails = "Request_x0020_Details";
        string FieldDueDate = "Request_x0020_Due_x0020_Date";
        string FieldAuthor = "Author";
        string FieldStatus = "Request_x0020_Status";

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
                item => item[FieldId],
                item => item[FieldDetails],
                item => item[FieldDueDate],
                item => item[FieldAuthor],
                item => item[FieldStatus]));
            clientContext.ExecuteQuery();

            ViewBag.Username = User.Identity.Name;
            ViewBag.List = spList;
            ViewBag.ListItems = spListItems;
            return View();
        }

        [HttpGet]
        public ActionResult NewItem()
        {
            ViewBag.Username = User.Identity.Name;
            return View();
        }

        [HttpPost]
        public RedirectResult NewItem(string details, DateTime dueDate)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);

            var user = clientContext.Web.EnsureUser(User.Identity.Name);
            clientContext.Load(user);
            clientContext.ExecuteQuery();
            FieldUserValue userValue = new FieldUserValue();
            userValue.LookupId = user.Id;

            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            var info = new ListItemCreationInformation();
            var item = spList.AddItem(info);
            item[FieldDetails] = details;
            item[FieldDueDate] = dueDate;
            item[FieldAuthor] = userValue;
            item.Update();
            clientContext.ExecuteQuery();

            return Redirect("Items");
        }

        [HttpGet]
        public ActionResult DeleteItem(int id)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);
            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            ListItem spListItem = spList.GetItemById(id);

            spListItem.DeleteObject();
            clientContext.ExecuteQuery();

            return RedirectToAction("Items");
        }

        [HttpGet]
        public ActionResult EditItem(int id)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);
            List spList = clientContext.Web.Lists.GetByTitle(ListName);

            clientContext.Load(spList);
            ListItem spListItem = spList.GetItemById(id);
            clientContext.Load(spListItem);
            clientContext.ExecuteQuery();
            
            string Request_ID = (String)spListItem[FieldId];
            string Request_Details = (String)spListItem[FieldDetails];
            string Request_Status = (String)spListItem[FieldStatus];
            var dueDate = spListItem[FieldDueDate];
            string Request_Due_Date = dueDate != null ? ((DateTime)dueDate).ToShortDateString() : "";
            FieldUserValue Author = (FieldUserValue)spListItem[FieldAuthor];
            string Request_By = Author.LookupValue;
            
            ViewBag.Request_ID = Request_ID;
            ViewBag.Request_Details = Request_Details;
            ViewBag.Request_Status = Request_Status;
            ViewBag.Request_Due_Date = Request_Due_Date;
            ViewBag.Request_By = Request_By;
            ViewBag.ID = id;

            // Getting choice fields from Request Status column 
            FieldChoice choiceField = clientContext.CastTo<FieldChoice>(spList.Fields.GetByInternalNameOrTitle(FieldStatus));
            clientContext.Load(choiceField);
            clientContext.ExecuteQuery();
            ViewBag.Request_Status_Choices = choiceField.Choices;

            return View();
        }

        [HttpPost]
        public ActionResult EditItem(int id, String details, DateTime dueDate, String status)
        {
            ClientContext clientContext = new ClientContext(SiteUrl);
            List spList = clientContext.Web.Lists.GetByTitle(ListName);
            ListItem spListItem = spList.GetItemById(id);

            spListItem[FieldDetails] = details;
            spListItem[FieldDueDate] = dueDate;
            spListItem[FieldStatus] = status;

            spListItem.Update();
            clientContext.ExecuteQuery();

            return RedirectToAction("Items");
        }
    }
}