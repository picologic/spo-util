using SPOUtil.ManagedMetadata;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using CommandLine;
using SPOUtil.Lists;

namespace SPOUtil
{
    class Program
    {
        static int Main(string[] args) {
            return CommandLine.Parser.Default.ParseArguments<
                MMSImportOptions, 
                MMSExportOptions,
                MMSClearOptions,
                ListImportOptions,
                ListUpdateOptions,
                ListClearOptions>(args)
                .MapResult(
                    (MMSImportOptions opts) => new ManagedMetadataOperations().ImportTerms(opts),
                    (MMSExportOptions opts) => new ManagedMetadataOperations().ExportTerms(opts),
                    (MMSClearOptions opts) => new ManagedMetadataOperations().ClearTerms(opts),
                    (ListImportOptions opts) => new ListOperations().ImportListData(opts),
                    (ListUpdateOptions opts) => new ListOperations().UpdateListData(opts),
                    (ListClearOptions opts) => new ListOperations().ClearItems(opts),
                    errs => 1);
        }
    }
}
