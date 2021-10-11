using System;
using System.Collections;
using System.ComponentModel;
using System.Configuration.Install;
using System.IO;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace ActionBIToolkit.InstallerActions
{
    [RunInstaller(true)]
    public class ExternalToolManifestInstaller : Installer
    {
        private static readonly string[] PowerShellProbePaths = new[] {
            "C:\\Program Files\\PowerShell\\7\\pwsh.exe",
            "C:\\Program Files\\PowerShell\\6\\pwsh.exe",
            "C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe"
        };

        public override void Commit(IDictionary savedState)
        {
            var manifestFile = new FileInfo(@"C:\Program Files (x86)\Common Files\Microsoft Shared\Power BI Desktop\External Tools\Action-BI-Toolkit.pbitool.json");
            if (manifestFile.Exists)
            {
                var psPath = PowerShellProbePaths.FirstOrDefault(File.Exists);
                if (psPath != null) // TODO Fail install?
                {
                    var manifestJson = JObject.Parse(File.ReadAllText(manifestFile.FullName));
                    manifestJson["path"] = psPath;
                    // TODO Set version number from installer version?

                    using (var manifestWriter = new Newtonsoft.Json.JsonTextWriter(File.CreateText(manifestFile.FullName)))
                    {
                        manifestWriter.Formatting = Newtonsoft.Json.Formatting.Indented;
                        manifestJson.WriteTo(manifestWriter);
                    }
                }
            }

            base.Commit(savedState);
        }


    }
}
