using Microsoft.SharePoint.Client;
using Microsoft.SharePoint.Client.Taxonomy;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SPOUtil.ManagedMetadata
{
    public class MetadataManager
    {
        private readonly ClientContext _ctx;
        private readonly TermStore _ts;

        public MetadataManager(ClientContext ctx, string termStore) {
            var session = TaxonomySession.GetTaxonomySession(ctx);
            _ctx = ctx;

            if (string.IsNullOrWhiteSpace(termStore)) {
                _ts = session.GetDefaultSiteCollectionTermStore();
                ctx.Load(_ts);
                ctx.ExecuteQuery();
            } else {
                var stores = session.TermStores;
                ctx.Load(stores);
                ctx.ExecuteQuery();
                _ts = stores.FirstOrDefault(x => x.Name == termStore);
            }

            if (_ts == null) {
                throw new Exception("Could not access term store");
            }
        }

        public void ImportTerms(string termGroup, string termSet, IEnumerable<TermData> terms) {
            var set = getTermSet(termGroup, termSet);

            var termNames = loadAllTermNamesInSet(set);
            foreach (var term in terms) {
                Console.Write("Importing term '{0}' ... ", term.Term);

                ClientResult<string> normalized = Term.NormalizeName(_ctx, term.Term);
                _ctx.ExecuteQuery();
                string normalizedTermName = normalized.Value;

                if (termNames.Contains(normalizedTermName.ToLower())) {
                    Console.WriteLine("SKIPPING Done");
                    continue;
                }

                Console.Write("CREATING ");
                importTerm(set, term);
                termNames.Add(normalizedTermName.ToLower());   // cache so we don't have to re-load

                Console.WriteLine("Done");
            }
        }

        public void ClearTerms(string termGroup, string termSet) {
            var set = getTermSet(termGroup, termSet);
            var terms = set.GetAllTerms();
            _ctx.Load(terms);
            _ctx.ExecuteQuery();

            foreach (var term in terms) {
                Console.Write("Deleting {0} ... ", term.Name);
                term.DeleteObject();
                _ctx.ExecuteQuery();
                Console.WriteLine("Done");
            }
        }

        private TermSet getTermSet(string termGroup, string termSet) {
            // get group
            var groups = _ts.Groups;
            _ctx.Load(groups);
            _ctx.ExecuteQuery();

            var group = groups.FirstOrDefault(x => x.Name == termGroup);
            if (group == null) {
                throw new Exception("Could not find term group " + termGroup);
            }
            var sets = group.TermSets;
            _ctx.Load(sets);
            _ctx.ExecuteQuery();

            // get set
            var set = sets.FirstOrDefault(x => x.Name == termSet);
            if (set == null) {
                throw new Exception("Could not find term set " + termSet);
            }

            return set;
        }

        private IList<string> loadAllTermNamesInSet(TermSet set) {
            var terms = set.GetAllTerms();
            _ctx.Load(terms);
            _ctx.ExecuteQuery();

            return terms.Select(x => x.Name.ToLower()).ToList();
        }

        private void importTerm(TermSet set, TermData term) {
            var newTerm = set.CreateTerm(term.Term, 1033, Guid.NewGuid());
            foreach (var p in term.Properties) {
                newTerm.SetCustomProperty(p.Key, p.Value);
            }
            _ctx.Load(newTerm);
            _ctx.ExecuteQuery();
        }
    }
}
