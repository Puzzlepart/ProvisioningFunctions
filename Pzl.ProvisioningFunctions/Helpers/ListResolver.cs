using System;
using System.Collections.Generic;
using Microsoft.SharePoint.Client;
using OfficeDevPnP.Core.Framework.Provisioning.Model;
using OfficeDevPnP.Core.Framework.Provisioning.ObjectHandlers;
using OfficeDevPnP.Core.Utilities;

namespace Pzl.ProvisioningFunctions.Helpers
{
    public class ListResolver
    {
        private readonly IDictionary<string, object> _properties;
        private readonly CanvasControl _control;
        private readonly Web _web;

        public ListResolver(CanvasControl control, Web web)
        {
            _control = control;
            _web = web;
            _properties = JsonUtility.Deserialize<Dictionary<string, dynamic>>(control.JsonControlData);
        }

        public void Process()
        {
            var list = GetList(_web);
            if (list == null) return;
            ClientObjectExtensions.EnsureProperties<List>(list, l => l.Id, l => l.RootFolder, l => l.RootFolder.Name);

            SetProperty("selectedListId", list.Id);
            SetProperty("selectedListUrl", list.RootFolder.Name);
            _control.JsonControlData = JsonUtility.Serialize<IDictionary<string, object>>(_properties);
        }

        private List GetList(Web web)
        {
            var listUrlProperty = GetProperty("selectedListUrl") as string;
            if (!string.IsNullOrWhiteSpace(listUrlProperty))
                return web.GetList(listUrlProperty);

            var listIdProperty = GetProperty("selectedListId") as string;
            Guid listId;
            if (TryParseGuidProperty(listIdProperty, out listId)) return web.Lists.GetById(listId);

            var listDisplayName = GetProperty("listTitle") as string;
            if (!string.IsNullOrWhiteSpace(listDisplayName)) return web.GetListByTitle(listDisplayName);

            return null;
        }

        private bool TryParseGuidProperty(string guid, out Guid id)
        {
            if (!string.IsNullOrWhiteSpace(guid) && Guid.TryParse(guid, out id) && !id.Equals(Guid.Empty)) return true;
            id = Guid.Empty;
            return false;
        }

        private object GetProperty(string name)
        {
            object value;
            return _properties.TryGetValue(name, out value) ? value : null;
        }

        private void SetProperty(string name, object value)
        {
            _properties[name] = value;
        }
    }
}