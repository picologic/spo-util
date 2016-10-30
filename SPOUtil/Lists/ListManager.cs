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
            var fields = _list.Fields;
            _ctx.Load(fields);
            _ctx.ExecuteQuery();

            var fieldMap = data.ColumnNames.Where(x => !string.IsNullOrEmpty(x)).ToDictionary(x => x, x => {
                var fieldTitle = x;
                if (mapping != null && mapping.ContainsKey(fieldTitle)) {
                    fieldTitle = mapping[fieldTitle];
                }

                var field = fields.Where(f => f.Title == fieldTitle || f.InternalName == fieldTitle).FirstOrDefault();
                if (field == null)
                    return "";

                return field.InternalName;
            });

            foreach (var row in data.Rows) {
                var itemCreateInfo = new ListItemCreationInformation();
                var item = _list.AddItem(itemCreateInfo);

                foreach (var colname in fieldMap.Keys) {
                    var fieldName = fieldMap[colname];
                    if (fieldName == "")    // if empty, it doesn't exist in target list so skip
                        continue;

                    var val = row.GetValueOrEmpty(colname);
                    if (!string.IsNullOrWhiteSpace(val)) {
                        // check if we need to convert value
                        var field = fields.Where(x => x.InternalName == fieldName).First();
                        switch (field.FieldTypeKind) {
                            case FieldType.DateTime:
                                DateTime dtVal;
                                if (!DateTime.TryParse(val, out dtVal)) {
                                    double dblVal;
                                    if (double.TryParse(val, out dblVal)) {
                                        dtVal = DateTime.FromOADate(dblVal);
                                    } else {
                                        throw new Exception("Could not parse " + val + " as DateTime");
                                    }
                                }

                                item[fieldName] = dtVal;
                                break;
                            case FieldType.URL:
                                if (!val.StartsWith("http://", StringComparison.InvariantCultureIgnoreCase) &&
                                    !val.StartsWith("https://", StringComparison.InvariantCultureIgnoreCase)) {
                                    val = "http://" + val;
                                }
                                item[fieldName] = val;
                                break;
                            default:
                                item[fieldName] = val;
                                break;
                        }
                        
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
