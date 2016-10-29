using CommandLine;
using DataAccess;
using System;
using System.Collections.Generic;

namespace SPOUtil.Lists
{
    public class ListOperations : OperationBase
    {
        public int ImportListData(ListImportOptions opts) {
            try {
                // parse csv
                var data = DataTable.New.ReadCsv(opts.FilePath);

                // load mapping
                IDictionary<string, string> mapping = null;
                if (string.IsNullOrWhiteSpace(opts.MappingFilePath) == false) {
                    mapping = new Dictionary<string, string>();
                    // load it
                }

                // initialize context
                WithContext(opts, ctx => {
                    var mgr = new ListManager(ctx, opts.List);
                    mgr.ImportItems(data, mapping);
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

    [Verb("import-list", HelpText = "Import records from a csv file into a SharePoint list")]
    public class ListImportOptions : SPOOptions
    {
        [Option('f', "file", Required = true, HelpText = "Data file containing data to import.")]
        public string FilePath { get; set; }

        [Option('l', "list", Required = true, HelpText = "List to import to.")]
        public string List { get; set; }

        [Option('m', "map", Required = false, HelpText = "File containing CSV to List column mapping. If not provided assume 1:1 match")]
        public string MappingFilePath { get; set; }
    }

    [Verb("clear-list", HelpText = "Remove all items from a SharePoint list.")]
    public class ListClearOptions : SPOOptions
    {
        [Option('l', "list", Required = false, HelpText = "List to remove items from")]
        public string List { get; set; }
    }
}
