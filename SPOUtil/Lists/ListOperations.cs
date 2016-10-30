using CommandLine;
using DataAccess;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace SPOUtil.Lists
{
    public class ListOperations : OperationBase
    {
        public int ImportListData(ListImportOptions opts) {
            return doImport(opts, (mgr, data, mapping, defaults) => mgr.ImportItems(data, mapping, defaults));
        }

        public int UpdateListData(ListUpdateOptions opts) {
            return doImport(opts, (mgr, data, mapping, defaults) => mgr.UpdateItems(data, mapping, defaults, opts.IDField));
        }

        private int doImport(ListImportOptions opts, Action<ListManager, MutableDataTable, IDictionary<string, string>, IDictionary<string, string>> importAction) {
            try {
                // parse data
                MutableDataTable data;
                if (opts.FilePath.EndsWith(".xlsx")) {
                    data = DataTable.New.ReadExcel(opts.FilePath);
                } else {
                    data = DataTable.New.ReadCsv(opts.FilePath);
                }

                Func<string, IDictionary<string, string>> loadSuppliment = (path) => {
                    IDictionary<string, string> dic = null;
                    if (string.IsNullOrWhiteSpace(path) == false) {
                        // load it
                        var lines = File.ReadAllLines(path);
                        dic = lines.Where(x => !string.IsNullOrWhiteSpace(x))
                                   .ToDictionary(x => x.Split(':')[0], x => x.Split(':')[1]);
                    }
                    return dic;
                };

                // load mapping and defaults
                var mapping = loadSuppliment(opts.MappingFilePath);
                var defaults = loadSuppliment(opts.DefaultFilePath);

                // initialize context
                WithContext(opts, ctx => {
                    var mgr = new ListManager(ctx, opts.List);
                    importAction(mgr, data, mapping, defaults);
                });

            } catch (Exception ex) {
                Console.WriteLine("An error occurred");
                Console.WriteLine(ex.Message);
                return 1;
            }
            return 0;
        }

        public int ClearItems(ListClearOptions opts) {
            try {
                // initialize context
                WithContext(opts, ctx => {
                    var mgr = new ListManager(ctx, opts.List);
                    mgr.DeleteAllItems();
                });

            } catch (Exception ex) {
                Console.WriteLine("An error occurred");
                Console.WriteLine(ex.Message);
                return 1;
            }

            return 0;
        }
    }

    [Verb("import-list", HelpText = "Import records from a csv or xslx file into a SharePoint list")]
    public class ListImportOptions : ListOptions
    {
        [Option('f', "file", Required = true, HelpText = "Data file containing data to import.")]
        public string FilePath { get; set; }

        [Option('m', "map", Required = false, HelpText = "File containing CSV to List column mapping. If not provided assume 1:1 match")]
        public string MappingFilePath { get; set; }

        [Option('d', "defaults", Required = false, HelpText = "File containing default values for columns not contained in data file")]
        public string DefaultFilePath { get; set; }
    }

    [Verb("update-list", HelpText = "Update items in a SharePoint list from a csv or xlsx file")]
    public class ListUpdateOptions : ListImportOptions
    {
        [Option('i', "idfield", Required = false, HelpText = "Field to use as the list item ID. If not provided ID or Title will be used by default")]
        public string IDField { get; set; }
    }

    [Verb("clear-list", HelpText = "Remove all items from a SharePoint list.")]
    public class ListClearOptions : ListOptions
    {
        
    }

    public abstract class ListOptions : SPOOptions
    {
        [Option('l', "list", Required = false, HelpText = "List to target")]
        public string List { get; set; }
    }
}
