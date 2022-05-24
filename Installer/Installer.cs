using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration.Install;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.RevitAddIns;

namespace Installer
{
    [RunInstaller(true)]
    public partial class Installer : System.Configuration.Install.Installer
    {
        public Installer()
        {
            InitializeComponent();
        }

        public override void Install(IDictionary stateSaver)
        {
            base.Install(stateSaver);

            WriteRevitAddin();
        }

        public override void Uninstall(IDictionary savedState)
        {
            base.Uninstall(savedState);

            DeleteRevitAddin();
        }

        private void WriteRevitAddin()
        {
            Guid guid = new Guid("3fc9040b-1de7-4a7f-9f3a-25d9c51b7217");
            string assembly = GetAssembly();
            string fullClassName = "BBI.JD.CrtlApplication";
            string vendorId = "JDS";
            string vendorDescription = "Juan Daniel SANTANA";

            foreach (string path in GetRevitVersionsPath())
            {
                string pathAddin = Path.Combine(path, "CenterGravity.addin");

                RevitAddInManifest manifest = File.Exists(pathAddin) ? AddInManifestUtility.GetRevitAddInManifest(pathAddin) : new RevitAddInManifest();

                RevitAddInApplication app = manifest.AddInApplications.FirstOrDefault(x => x.AddInId == guid);

                if (app == null)
                {
                    app = new RevitAddInApplication("CenterGravity", assembly, guid, fullClassName, vendorId);
                    app.VendorDescription = vendorDescription;

                    manifest.AddInApplications.Add(app);
                }
                else
                {
                    app.Assembly = assembly;
                    app.FullClassName = fullClassName;
                }

                if (manifest.Name == null)
                {
                    manifest.SaveAs(pathAddin);
                }
                else
                {
                    manifest.Save();
                }
            }
        }

        private void DeleteRevitAddin()
        {
            foreach (string path in GetRevitVersionsPath())
            {
                string pathAddin = Path.Combine(path, "CenterGravity.addin");

                if (File.Exists(pathAddin))
                {
                    File.Delete(pathAddin);
                }
            }
        }

        private string GetAssembly()
        {
            string pathDir = Context.Parameters["targetdir"];

            pathDir = pathDir.Remove(pathDir.Length - 1, 1);

            return pathDir + "CenterGravity.dll";
        }

        private List<string> GetRevitVersionsPath()
        {
            List<string> paths = new List<string>();

            RevitProduct product = RevitProductUtility.GetAllInstalledRevitProducts()
                .FirstOrDefault(x => x.Version != RevitVersion.Unknown);

            if (product != null)
            {
                DirectoryInfo parent = Directory.GetParent(product.AllUsersAddInFolder);

                int ver;

                foreach (DirectoryInfo version in parent.EnumerateDirectories())
                {
                    if (int.TryParse(version.Name, out ver))
                    {
                        // The plugin was compiled for versions equal to 2019 and above
                        if (ver >= 2019)
                        {
                            paths.Add(version.FullName);
                        }
                    }
                }
            }

            return paths;
        }
    }
}
