using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using Microsoft.Win32;

namespace OpenExplorer
{
    public partial class MyRibbon
    {
        public OneDriveInfo []infos;
        private void MyRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            infos = OneDriveInfo.GetOneDriveInfos().ToArray();
        }

        private void OpenInExplorerButton_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            var doc = app.ActiveDocument;
            var path = doc.FullName;
            var uri = new Uri(path);
            if (uri.IsFile)
            {
                OpenFolder(path);
                return;
            }
            else {
                foreach (var info in infos)
                {
                    if (info.IsMyLink(path)) {
                        var local_path = info.GetLocalPath(path);
                        OpenFolder(local_path);
                        return;
                    }
                }
            }
        }

        private static void OpenFolder(string path)
        {
            if (System.IO.File.Exists(path))
            {
                var p = new System.Diagnostics.Process();
                p.StartInfo.FileName = "explorer.exe";
                p.StartInfo.Arguments = string.Format("/select,\"{0}\"", path);
                p.Start();
                p.Close();
                return;
            }

            var dir = System.IO.Path.GetDirectoryName(path);
            if (System.IO.Directory.Exists(dir))
            {
                var p = new System.Diagnostics.Process();
                p.StartInfo.FileName = "explorer.exe";
                p.StartInfo.Arguments = dir;
                p.Start();
                p.Close();
                return;
            }
        }
    }
}
