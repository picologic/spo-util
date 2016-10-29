using CommandLine;
using Microsoft.SharePoint.Client;
using SPOUtil.Extensions;
using System;
using System.Net;

namespace SPOUtil
{
    public abstract class OperationBase
    {
        protected void WithContext(SPOOptions opts, Action<ClientContext> action) {
            using (var ctx = new ClientContext(opts.Url)) {
                ctx.AuthenticationMode = ClientAuthenticationMode.Default;
                ctx.Credentials = getCredentials(opts.Username, opts.Password);

                action(ctx);
            }
        }

        private ICredentials getCredentials(string username, string password) {
            return new SharePointOnlineCredentials(username, password.ToSecureString());
        }
    }

    public abstract class SPOOptions
    {
        [Option('r', "url", Required = true, HelpText = "Url to the site or tenant")]
        public string Url { get; set; }

        [Option('u', "username", Required = true, HelpText = "Username to authenticate")]
        public string Username { get; set; }

        [Option('p', "password", Required = true, HelpText = "Password to authenticate")]
        public string Password { get; set; }
    }
}
