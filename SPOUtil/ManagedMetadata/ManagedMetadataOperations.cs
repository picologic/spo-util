using CommandLine;
using DataAccess;
using System;
using System.Linq;
using System.Collections.Generic;

namespace SPOUtil.ManagedMetadata
{
    public class ManagedMetadataOperations : OperationBase
    {
        public int ImportTerms(MMSImportOptions opts) {
            try {
                // parse csv
                var data = DataTable.New.ReadCsv(opts.FilePath);
                var terms = new List<TermData>();
                foreach (var row in data.Rows) {
                    var term = new TermData {
                        Term = row.GetValueOrEmpty("Term")
                    };
                    foreach (var col in data.ColumnNames) {
                        if (col.StartsWith("Prop_")) {
                            var prop = col.Substring(5);
                            term.Properties[prop] = row.GetValueOrEmpty(col);
                        }
                    }
                    terms.Add(term);
                }

                WithContext(opts, ctx => {
                    // import terms
                    var mgr = new MetadataManager(ctx, opts.TermStore);
                    mgr.ImportTerms(opts.TermGroup, opts.TermSet, terms);
                });

            } catch (Exception ex) {
                Console.WriteLine("An error occurred");
                Console.WriteLine(ex.Message);
                return 1;
            }
            return 0;
        }

        public int ExportTerms(MMSExportOptions opts) {
            try {
                IDictionary<string, string[]> terms = null;
                WithContext(opts, ctx => {
                    var mgr = new MetadataManager(ctx, opts.TermStore);
                    terms = mgr.ExportTerms(opts.TermGroup, opts.TermSet);
                });

                var flat = terms.SelectMany(kvp => {
                    return kvp.Value.Select(t => {
                        return new {
                            TermSet = kvp.Key,
                            Term = t
                        };
                    });
                });

                DataTable.New.FromEnumerable(flat).SaveCSV(opts.FilePath);
                
            } catch (Exception ex) {
                Console.WriteLine("An error occurred");
                Console.WriteLine(ex.Message);
                return 1;
            }
            return 0;
        }

        public int ClearTerms(MMSClearOptions opts) {
            try {
                WithContext(opts, ctx => {
                    var mgr = new MetadataManager(ctx, opts.TermStore);
                    mgr.ClearTerms(opts.TermGroup, opts.TermSet);
                });
            } catch (Exception ex) {
                Console.WriteLine("An error occurred");
                Console.WriteLine(ex.Message);
                return 1;
            }

            return 0;
        }
    }

    [Verb("import-terms", HelpText = "Import terms from a csv file into a term set in the term store")]
    public class MMSImportOptions : SPOOptions
    {
        [Option('f', "file", Required = true, HelpText = "Data file containing data to import.")]
        public string FilePath { get; set; }

        [Option('t', "termStore", Required = false, HelpText = "Term store to import to.  If empty the default site collection term store will be used")]
        public string TermStore { get; set; }

        [Option('g', "termGroup", Required = true, HelpText = "Term group containing the term set to import to.")]
        public string TermGroup { get; set; }

        [Option('s', "termSet", Required = true, HelpText = "Term set to import terms into.")]
        public string TermSet { get; set; }
    }

    [Verb("clear-terms", HelpText = "Remove all terms from a term set in the term store.")]
    public class MMSClearOptions : SPOOptions
    {
        [Option('t', "termStore", Required = false, HelpText = "Term store to modify.  If empty the default site collection term store will be used")]
        public string TermStore { get; set; }

        [Option('g', "termGroup", Required = true, HelpText = "Term group containing the term set to clear.")]
        public string TermGroup { get; set; }

        [Option('s', "termSet", Required = true, HelpText = "Term set to clear.")]
        public string TermSet { get; set; }
    }

    [Verb("export-terms", HelpText = "Export all terms to a csv file")]
    public class MMSExportOptions : SPOOptions
    {
        [Option('f', "file", Required = true, HelpText = "File to write output to")]
        public string FilePath { get; set; }

        [Option('t', "termStore", Required = false, HelpText = "Term store to export from. If empty the default site collection term store will be used")]
        public string TermStore { get; set; }

        [Option('g', "termGroup", Required = true, HelpText = "Term group containing term sets to export")]
        public string TermGroup { get; set; }

        [Option('s', "termSet", Required = false, HelpText = "Term set to export. If empty all term sets in the term group will be exported")]
        public string TermSet { get; set; }
    }
}
