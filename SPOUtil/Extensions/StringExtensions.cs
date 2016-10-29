using System.Security;

namespace SPOUtil.Extensions
{
    public static class StringExtensions
    {
        public static SecureString ToSecureString(this string str) {
            var secureStr = new SecureString();
            foreach (var c in str) {
                secureStr.AppendChar(c);
            }
            secureStr.MakeReadOnly();
            return secureStr;
        }
    }
}
