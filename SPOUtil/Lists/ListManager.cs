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

        public void ImportItems(MutableDataTable data, IDictionary<string, string> mapping = null, IDictionary<string, string> defaults = null) {
            var fields = _list.Fields;
            _ctx.Load(fields);
            _ctx.ExecuteQuery();

            var fieldMap = mapColumnsToFields(data, mapping, defaults, fields);

            var idx = 0;
            foreach (var row in data.Rows) {
                idx++;
                Console.WriteLine("Importing {0} of {1} records", idx, data.NumRows);

                var itemCreateInfo = new ListItemCreationInformation();
                var item = _list.AddItem(itemCreateInfo);

                populateItem(item, row, fieldMap, defaults, fields);

                item.Update();
                _ctx.ExecuteQuery();
            }
        }

        public void UpdateItems(MutableDataTable data, IDictionary<string, string> mapping = null, IDictionary<string, string> defaults = null, string idField = "") {
            var fields = _list.Fields;
            _ctx.Load(fields);
            _ctx.ExecuteQuery();

            var fieldMap = mapColumnsToFields(data, mapping, defaults, fields);

            if (idField == "") {
                var cols = fieldMap.Where(x => x.Value == "ID" || x.Value == "Title").OrderBy(x => x.Value).ToList();
                if (cols.Count == 0)
                    throw new Exception("Data source does not contain an ID or Title field and no idField argument was provided.");

                idField = cols[0].Key;
            }

            var idx = 0;
            foreach (var row in data.Rows) {
                var idval = row[idField];

                idx++;
                Console.WriteLine("Updating {0} of {1} records with ID {2}", idx, data.NumRows, idval);

                var idFieldInfo = fields.Where(x => x.InternalName == fieldMap[idField]).First();
                var idquery = string.Format("<FieldRef Name='{0}'/><Value Type='{1}'>{2}</Value>", fieldMap[idField], idFieldInfo.FieldTypeKind, idval);

                var query = new CamlQuery();
                query.ViewXml = "<View><Query><Where><Eq>" + idquery + "</Eq></Where></Query></View>";

                var items = _list.GetItems(query);
                _ctx.Load(items);
                _ctx.ExecuteQuery();
                if (items.Count == 0) {
                    Console.WriteLine("SKIPPING - could not find item in list");
                    continue;
                }

                var item = items[0];

                populateItem(item, row, fieldMap, defaults, fields);

                item.Update();
                _ctx.ExecuteQuery();
            }
        }

        private IDictionary<string, string> mapColumnsToFields(MutableDataTable data, IDictionary<string, string> mapping, IDictionary<string, string> defaults, FieldCollection fields) {
            var columnNames = data.ColumnNames.Where(x => !string.IsNullOrEmpty(x)).ToList();
            if (defaults != null) {
                columnNames.AddRange(defaults.Keys.ToList());
            }

            var fieldMap = columnNames.ToDictionary(x => x, x => {
                var fieldTitle = x;
                if (mapping != null && mapping.ContainsKey(fieldTitle)) {
                    fieldTitle = mapping[fieldTitle];
                }

                var field = fields.Where(f => f.Title == fieldTitle || f.InternalName == fieldTitle).FirstOrDefault();
                if (field == null)
                    return "";

                return field.InternalName;
            });

            return fieldMap;
        }

        private void populateItem(ListItem item, Row row, IDictionary<string, string> fieldMap, IDictionary<string, string> defaults, FieldCollection fields) {
            // helper
            Action<string, string> setField = (fieldName, val) => {
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
                        case FieldType.Lookup:
                            FieldLookup lookup = (FieldLookup)field;
                            item[fieldName] = getLookupValue(val, lookup);
                            break;
                        case FieldType.Currency:
                            double dbVal;
                            if (double.TryParse(val, out dbVal)) {
                                item[fieldName] = dbVal;
                            } else {
                                Console.WriteLine("Could not convert {0} to currency", val);
                            }
                            break;
                        case FieldType.Text:
                            FieldText tfield = (FieldText)field;
                            if (val.Length > tfield.MaxLength) {
                                val = val.Substring(0, tfield.MaxLength);
                                Console.WriteLine("Truncating value for {0}", tfield.Title);
                            }
                            item[fieldName] = val;
                            break;
                        default:
                            item[fieldName] = val;
                            break;
                    }
                }
            };
            
            foreach (var colname in fieldMap.Keys) {
                var fieldName = fieldMap[colname];
                if (fieldName == "")    // if empty, it doesn't exist in target list so skip
                    continue;

                var val = "";
                if (row.ColumnNames.Contains(colname))
                    val = row.GetValueOrEmpty(colname);

                if (string.IsNullOrWhiteSpace(val) && defaults.ContainsKey(colname))
                    val = defaults[colname];

                setField(fieldName, val);
            }
        }

        private FieldLookupValue getLookupValue(string value, FieldLookup field) {
            var lookupList = _ctx.Web.Lists.GetById(new Guid(field.LookupList));
            var query = new CamlQuery();
            query.ViewXml = string.Format("<View><Query><Where><Eq><FieldRef Name='{0}'/><Value Type='Text'>{1}</Value></Eq></Where></Query></View>", field.LookupField, value);

            var items = lookupList.GetItems(query);
            _ctx.Load(items, i => i.Include(li => li.Id));
            _ctx.ExecuteQuery();

            if (items.Count == 0) {
                return null;
            }

            var lv = new FieldLookupValue();
            lv.LookupId = items[0].Id;
            return lv;
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
