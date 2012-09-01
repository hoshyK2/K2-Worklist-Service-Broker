using System;
using System.Collections.Generic;
using System.Text;
using System.Xml;
using System.Data;
using SourceCode.SmartObjects.Services.ServiceSDK.Types;
using SourceCode.SmartObjects.Services.ServiceSDK.Objects;
using SourceCode.Hosting.Client;
using SourceCode.Workflow.Client;
using SourceCode.Hosting.Client.BaseAPI;
using System.Web;

namespace SourceCode.SmartObjects.Services.WorklistService
{
    public class BasicWorklistItem
    {
        #region private data members

        private string _connectionString;
        private string _connectionstringImpersonate;

        #endregion

        #region ctor(s)

        internal BasicWorklistItem() { }

        internal BasicWorklistItem(string connectionString, string connectionStringImpersonate)
        {
            _connectionString = connectionString;
            _connectionstringImpersonate = connectionStringImpersonate;
        }

        #endregion

        /// <summary>
        /// Describes the functionality of the service object
        /// </summary>
        /// <returns>Returns the service object definition.</returns>
        internal ServiceObject DescribeServiceObject()
        {
            ServiceObject so = new ServiceObject("BasicWorklistItem");
            so.Type = "BasicWorklistItem";
            so.Active = true;
            so.MetaData.DisplayName = "Basic Worklist Item";
            so.MetaData.Description = "Represents a Basic WorklistItem";
            so.Properties = DescribeProperties();
            so.Methods = DescribeMethods();
            return so;
        }

        private SourceCode.SmartObjects.Services.ServiceSDK.Objects.Properties DescribeProperties()
        {
            SourceCode.SmartObjects.Services.ServiceSDK.Objects.Properties properties = new SourceCode.SmartObjects.Services.ServiceSDK.Objects.Properties();
            // worklistitem properties
            properties.Add(new Property("AllocatedUser", "System.String", SoType.Text, new MetaData("Allocated User", "AllocatedUser")));
            properties.Add(new Property("Data", "System.String", SoType.Text, new MetaData("Data", "Data")));
			properties.Add(new Property("Link", "System.String", SoType.HyperLink, new MetaData("Link", "Link")));
            properties.Add(new Property("ID", "System.Int32", SoType.Autonumber, new MetaData("ID", "ID")));
            properties.Add(new Property("SerialNumber", "System.String", SoType.Text, new MetaData("SerialNumber", "SerialNumber")));
            properties.Add(new Property("Status", "System.String", SoType.Text, new MetaData("Status", "Status")));

            // activity instance destination properties
            properties.Add(new Property("ActivityID", "System.Int32", SoType.Number, new MetaData("Activity ID", "Activity ID")));
            properties.Add(new Property("ActivityInstanceID", "System.Int32", SoType.Number, new MetaData("Activity Instance ID", "Activity Instance ID")));
            properties.Add(new Property("ActivityInstanceDestinationID", "System.Int32", SoType.Number, new MetaData("Activity Instance Destination ID", "Activity Instance Destination ID")));
            properties.Add(new Property("ActivityName", "System.String", SoType.Text, new MetaData("Activity Name", "Activity Name")));
            properties.Add(new Property("Priority", "System.String", SoType.Text, new MetaData("Priority", "Priority")));
            properties.Add(new Property("StartDate", "System.DateTime", SoType.DateTime, new MetaData("StartDate", "StartDate")));

            // process instance properties
            properties.Add(new Property("ProcessInstanceID", "System.Int32", SoType.Number, new MetaData("Process Instance ID", "Process Instance ID")));
            properties.Add(new Property("ProcessFullName", "System.String", SoType.Text, new MetaData("Process Full Name", "Process Full Name")));
            properties.Add(new Property("ProcessName", "System.String", SoType.Text, new MetaData("Process Name", "Process Name")));
            properties.Add(new Property("Folio", "System.String", SoType.Text, new MetaData("Folio", "Folio")));

            // event instance properties
            properties.Add(new Property("EventInstanceName", "System.String", SoType.Text, new MetaData("Event Instance Name", "Event Instance Name")));

            // username property for impersonation
            properties.Add(new Property("UserName", "System.String", SoType.Text, new MetaData("User Name", "User Name")));
            return properties;
        }

        private SourceCode.SmartObjects.Services.ServiceSDK.Objects.Methods DescribeMethods()
        {
            SourceCode.SmartObjects.Services.ServiceSDK.Objects.Methods methods = new Methods();
            methods.Add(new Method("GetWorklistItems", MethodType.List, new MetaData("Get Worklist Items", "Returns a collection of worklist items."), GetRequiredProperties("GetWorklistItems"), GetMethodParameters("GetWorklistItems"), GetInputProperties("GetWorklistItems"), GetReturnProperties("GetWorklistItems")));
            methods.Add(new Method("LoadWorklistItem", MethodType.Read, new MetaData("Load Worklist Item", "Returns the specified worklist item."), GetRequiredProperties("LoadWorklistItem"), GetMethodParameters("LoadWorklistItem"), GetInputProperties("LoadWorklistItem"), GetReturnProperties("LoadWorklistItem")));
            return methods;
        }

        private InputProperties GetInputProperties(string method)
        {
            InputProperties properties = new InputProperties();

            switch (method)
            {
                case "GetWorklistItems":
                    {
                        #region properties
                        properties.Add(new Property("Status", "System.String", SoType.Text, new MetaData("Status", "Status")));
                        properties.Add(new Property("ActivityName", "System.String", SoType.Text, new MetaData("Activity Name", "Activity Name")));
                        properties.Add(new Property("ProcessName", "System.String", SoType.Text, new MetaData("Process Name", "Process Name")));
                        properties.Add(new Property("ProcessFullName", "System.String", SoType.Text, new MetaData("Process Full Name", "Process Full Name")));
                        properties.Add(new Property("Folio", "System.String", SoType.Text, new MetaData("Folio", "Folio")));
                        properties.Add(new Property("EventInstanceName", "System.String", SoType.Text, new MetaData("Event Instance Name", "Event Instance Name")));
                        properties.Add(new Property("Priority", "System.String", SoType.Text, new MetaData("Priority", "Priority")));
                        properties.Add(new Property("UserName", "System.String", SoType.Text, new MetaData("User Name", "User Name")));
                        break;
                        #endregion
                    }
                case "LoadWorklistItem":
                    {
                        properties.Add(new Property("SerialNumber", "System.String", SoType.Text, new MetaData("SerialNumber", "SerialNumber")));
                        break;
                    }
            }
            return properties;
        }

        private Validation GetRequiredProperties(string method)
        {
            RequiredProperties properties = new RequiredProperties();
            Validation validation = null;
            validation = new Validation();
            switch (method)
            {
                case "GetWorklistItems":
                    {
                        break;
                    }
                case "LoadWorklistItem":
                    {
                        properties.Add(new Property("SerialNumber", "System.String", SoType.Text, new MetaData("SerialNumber", "SerialNumber")));
                        break;
                    }
            }
            validation.RequiredProperties = properties;
            return validation;
        }

        private MethodParameters GetMethodParameters(string method)
        {
            MethodParameters parameters = new MethodParameters();
            switch (method)
            {
                case "GetWorklistItems":
                    {
                        break;
                    }
            }
            return parameters;
        }

        private ReturnProperties GetReturnProperties(string method)
        {
            ReturnProperties properties = new ReturnProperties();
            switch (method)
            {
                case "GetWorklistItems":
                case "LoadWorklistItem":
                    {
                        #region GetWorklistItems
                        // worklistitem properties
                        properties.Add(new Property("AllocatedUser", "System.String", SoType.Text, new MetaData("Allocated User", "AllocatedUser")));
                        properties.Add(new Property("Data", "System.String", SoType.Text, new MetaData("Data", "Data")));
						properties.Add(new Property("Link", "System.String", SoType.HyperLink, new MetaData("Link", "Link")));
                        properties.Add(new Property("ID", "System.Int32", SoType.Autonumber, new MetaData("ID", "ID")));
                        properties.Add(new Property("SerialNumber", "System.String", SoType.Text, new MetaData("SerialNumber", "SerialNumber")));
                        properties.Add(new Property("Status", "System.String", SoType.Text, new MetaData("Status", "Status")));

                        // activity instance destination properties
                        properties.Add(new Property("ActivityID", "System.Int32", SoType.Number, new MetaData("Activity ID", "Activity ID")));
                        properties.Add(new Property("ActivityInstanceID", "System.Int32", SoType.Number, new MetaData("Activity Instance ID", "Activity Instance ID")));
                        properties.Add(new Property("ActivityInstanceDestinationID", "System.Int32", SoType.Number, new MetaData("Activity Instance Destination ID", "Activity Instance Destination ID")));
                        properties.Add(new Property("ActivityName", "System.String", SoType.Text, new MetaData("Activity Name", "Activity Name")));
                        properties.Add(new Property("Priority", "System.String", SoType.Text, new MetaData("Priority", "Priority")));
                        properties.Add(new Property("StartDate", "System.DateTime", SoType.DateTime, new MetaData("StartDate", "StartDate")));

                        // process instance properties
                        properties.Add(new Property("ProcessInstanceID", "System.Int32", SoType.Number, new MetaData("Process Instance ID", "Process Instance ID")));
                        properties.Add(new Property("ProcessFullName", "System.String", SoType.Text, new MetaData("Process Full Name", "Process Full Name")));
                        properties.Add(new Property("ProcessName", "System.String", SoType.Text, new MetaData("Process Name", "Process Name")));
                        properties.Add(new Property("Folio", "System.String", SoType.Text, new MetaData("Folio", "Folio")));
                        properties.Add(new Property("EventInstanceName", "System.String", SoType.Text, new MetaData("Event Instance Name", "Event Instance Name")));
                        break;
                        #endregion
                    }
            }
            return properties;
        }

        internal DataTable LoadWorklistItem(string serialNumber)
        {
            ConnectionSetup connectSetup = new ConnectionSetup();
            connectSetup.ConnectionString = _connectionString;
            Connection cnn = new Connection();

            DataTable dt = GetResultTable();

            try
            {
                cnn.Open(connectSetup);
                WorklistItem item = cnn.OpenWorklistItem(serialNumber);

                DataRow row = dt.NewRow();
                row["AllocatedUser"] = item.AllocatedUser;
                row["Data"] = item.Data;
                row["ID"] = item.ID;
				row["Link"] = "<hyperlink><link>" + HttpUtility.HtmlEncode(item.Data) + "</link><display>Open</display></hyperlink>";
                row["SerialNumber"] = item.SerialNumber;
                row["Status"] = item.Status;
                row["ActivityID"] = item.ActivityInstanceDestination.ActID;
                row["ActivityInstanceID"] = item.ActivityInstanceDestination.ActInstID;
                row["ActivityName"] = item.ActivityInstanceDestination.Name;
                row["Priority"] = item.ActivityInstanceDestination.Priority;
                row["StartDate"] = item.EventInstance.StartDate;
                row["ActivityInstanceDestinationID"] = item.ActivityInstanceDestination.ID;
                row["ProcessInstanceID"] = item.ProcessInstance.ID;
                row["ProcessFullName"] = item.ProcessInstance.FullName;
                row["ProcessName"] = item.ProcessInstance.Name;
                row["Folio"] = item.ProcessInstance.Folio;
                row["EventInstanceName"] = item.EventInstance.Name;
                dt.Rows.Add(row);
                return dt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                cnn.Close();
                cnn.Dispose();
            }
        }

        public DataTable GetWorklistItems(Dictionary<string, object> properties, Dictionary<string, object> parameters)
        {
            bool impersonate = false;
            string impersonateUser = "";
            ConnectionSetup connectSetup = new ConnectionSetup();
            connectSetup.ConnectionString = _connectionString;

            if (properties.ContainsKey("UserName"))
            {
                if (!(string.IsNullOrEmpty(properties["UserName"].ToString())))
                {
                    connectSetup.ConnectionString = _connectionstringImpersonate;
                    impersonateUser = properties["UserName"].ToString();
                    impersonate = true;
                }
                else
                    connectSetup.ConnectionString = _connectionString;
            }

            WorklistCriteria criteria = null;
            if (properties.Count > 0)
                criteria = GetWorklistCriteria(properties);
           
            Connection cnn = new Connection();
            try
            {
                cnn.Open(connectSetup);
                if (impersonate)
                    cnn.ImpersonateUser(impersonateUser);

                DataTable dt = GetResultTable();

                Worklist worklist;
                if ((criteria != null) && (criteria.Filters.GetLength(0) > 0))
                    worklist = cnn.OpenWorklist(criteria);
                else
                    worklist = cnn.OpenWorklist();
                foreach (WorklistItem item in worklist)
                {
                    DataRow row = dt.NewRow();
                    row["AllocatedUser"] = item.AllocatedUser;
                    row["Data"] = item.Data;
                    row["ID"] = item.ID;
                    row["Link"] = "<hyperlink><link>" + HttpUtility.HtmlEncode(item.Data) + "</link><display>Open</display></hyperlink>";
                    row["SerialNumber"] = item.SerialNumber;
                    row["Status"] = item.Status;
                    row["ActivityID"] = item.ActivityInstanceDestination.ActID;
                    row["ActivityInstanceID"] = item.ActivityInstanceDestination.ActInstID;
                    row["ActivityName"] = item.ActivityInstanceDestination.Name;
                    row["Priority"] = item.ActivityInstanceDestination.Priority;
                    row["StartDate"] = item.EventInstance.StartDate;
                    row["ActivityInstanceDestinationID"] = item.ActivityInstanceDestination.ID;
                    row["ProcessInstanceID"] = item.ProcessInstance.ID;
                    row["ProcessFullName"] = item.ProcessInstance.FullName;
                    row["ProcessName"] = item.ProcessInstance.Name;
                    row["Folio"] = item.ProcessInstance.Folio;
                    row["EventInstanceName"] = item.EventInstance.Name;
                    dt.Rows.Add(row);
                }
                if (impersonate)
                    cnn.RevertUser();
                cnn.Close();
                cnn.Dispose();
                return dt;

            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private WorklistCriteria GetWorklistCriteria(Dictionary<string, object> properties)
        {
            WorklistCriteria criteria = new WorklistCriteria();

            foreach (KeyValuePair<string, object> property in properties)
            {
                switch (property.Key)
                {
                    case "Status":
                        criteria.AddFilterField(WCField.WorklistItemStatus, WCCompare.Equal, property.Value);
                        break;
                    case "ActivityName":
                        criteria.AddFilterField(WCField.ActivityName, WCCompare.Equal, property.Value);
                        break;
                    case "ProcessName":
                        criteria.AddFilterField(WCField.ProcessName, WCCompare.Equal, property.Value);
                        break;
                    case "ProcessFullName":
                        criteria.AddFilterField(WCField.ProcessFullName, WCCompare.Equal, property.Value);
                        break;
                    case "Folio":
                        criteria.AddFilterField(WCField.ProcessFolio, WCCompare.Equal, property.Value);
                        break;
                    case "EventInstanceName":
                        criteria.AddFilterField(WCField.EventName, WCCompare.Equal, property.Value);
                        break;
                    case "Priority":
                        criteria.AddFilterField(WCField.ActivityPriority, WCCompare.Equal, property.Value);
                        break;
                }
            }

            return criteria;
        }

        private DataTable GetResultTable()
        {
            DataTable result = new DataTable("BasicWorklistItem");
            result.Columns.Add("AllocatedUser", typeof(string));
            result.Columns.Add("Data", typeof(string));
			result.Columns.Add("Link", typeof(string));
            result.Columns.Add("ID", typeof(Int32));
            result.Columns.Add("SerialNumber", typeof(string));
            result.Columns.Add("Status", typeof(string));
            result.Columns.Add("ActivityID", typeof(Int32));
            result.Columns.Add("ActivityInstanceID", typeof(Int32));
            result.Columns.Add("ActivityName", typeof(string));
            result.Columns.Add("Priority", typeof(string));
            result.Columns.Add("StartDate", typeof(DateTime));
            result.Columns.Add("ActivityInstanceDestinationID", typeof(Int32));
            result.Columns.Add("ProcessInstanceID", typeof(Int32));
            result.Columns.Add("ProcessFullName", typeof(string));
            result.Columns.Add("ProcessName", typeof(string));
            result.Columns.Add("Folio", typeof(string));
            result.Columns.Add("EventInstanceName", typeof(string));

            return result;
        }
    }
}
