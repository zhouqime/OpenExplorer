using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Win32;

namespace OpenExplorer
{
    public class OneDriveInfo
    {
        public string UserFolder { get; set; }
        public string Root { get; set; }
        public string DocUri { get { return Root + "Documents/"; } }

        public bool IsMyLink(string path) {
            var uri = new Uri(path);
            return uri.ToString().StartsWith(DocUri);
        }

        public string GetLocalPath(string path) {
            var uri = new Uri(path);
            return uri.ToString().Replace(DocUri, UserFolder + "/").Replace('/','\\');
        }

        static string GetParentDirectory(Uri uri)
        {
            return uri.Scheme + "://" + uri.Host + string.Join("", new List<string>(uri.Segments).GetRange(0, uri.Segments.Length - 1));
        }

        public static IEnumerable<OneDriveInfo> GetOneDriveInfos() {
            var key = Registry.CurrentUser.OpenSubKey(@"SOFTWARE\Microsoft\OneDrive\Accounts");
            //key = key.OpenSubKey("Microsoft");
            //key = key.OpenSubKey("OneDrive");
            //key = key.OpenSubKey("Accounts");

            foreach (var name in key.GetSubKeyNames()) {
                var ac = key.OpenSubKey(name);
                var uri = ac.GetValue("ServiceEndpointUri").ToString();

                yield return new OneDriveInfo()
                {
                    UserFolder = ac.GetValue("UserFolder").ToString(),
                    Root = uri.Substring(0,uri.Length - 4)// 移除_api部分
                };
            }
        }
    }
}
