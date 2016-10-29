using DataAccess;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SPOUtil.Lists
{
    public class ListManager
    {
        private readonly ClientContext _ctx;
        private readonly List _list;

        public ListManager(ClientContext ctx, string list) {
            _ctx = ctx;

            // load list
            _list = ctx.Web.Lists.GetByTitle(list);
            ctx.Load(_list);
            ctx.ExecuteQuery();

            if (_list == null) {
                throw new Exception("Could not find list");
            }
        }

        public void ImportItems(MutableDataTable data, IDictionary<string, string> mapping = null) {
            foreach (var row in data.Rows) {
                var itemCreateInfo = new ListItemCreationInformation();
                var item = _list.AddItem(itemCreateInfo);

                foreach (var colname in data.ColumnNames) {
                    var fieldName = colname;
                    if (mapping != null && mapping.ContainsKey(colname))
                        fieldName = mapping[colname];

                    var val = row.GetValueOrEmpty(colname);
                    if (!string.IsNullOrWhiteSpace(val)) {
                        item[fieldName] = val;
                    }
                }

                item.Update();
                _ctx.ExecuteQuery();
            }
        }

        public void DeleteAllItems() {
            var itemCount = _list.ItemCount;
            ListItemCollectionPosition licp = null;
            while (true) {
                var query = new CamlQuery();
                query.ViewXml = "<View><ViewFields><FieldRef Name='Id'/></ViewFields><RowLimit>250</RowLimit></View>";
                query.ListItemCollectionPosition = licp;
                var items = _list.GetItems(query);
                _ctx.Load(items);
                _ctx.ExecuteQuery();

                Console.WriteLine("Deleting {0} of {1} items", items.Count, itemCount);
                itemCount -= items.Count;

                foreach (var item in items.ToList()) {
                    item.DeleteObject();
                }
                _ctx.ExecuteQuery();
                if (licp == null) {
                    Console.WriteLine("Finished clearing list");
                    break;
                }
            }
        }
    }
}
