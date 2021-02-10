using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace ZipTestApp
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string rootPath = AppDomain.CurrentDomain.BaseDirectory;

            string pointsFolder = Path.Combine(rootPath, "Points");
            if (!Directory.Exists(pointsFolder))
            {
                System.IO.Directory.CreateDirectory(pointsFolder);
            }

            string tempPath = Path.Combine(rootPath, @"Temp");
            if (!Directory.Exists(tempPath))
            {
                System.IO.Directory.CreateDirectory(tempPath);
            }

            string extractPath = Path.Combine(rootPath, @"KML");
            if (!Directory.Exists(extractPath))
            {
                System.IO.Directory.CreateDirectory(extractPath);
            }

            foreach (var file in Directory.EnumerateFiles(pointsFolder, "*.kmz"))
            {
                string fileName = Path.GetFileNameWithoutExtension(file);

                try
                {
                    System.IO.Compression.ZipFile.ExtractToDirectory(file, tempPath);
                }
                catch (Exception ex)
                {
                    textBox1.AppendText(ex.Message);
                    textBox1.AppendText(file);
                    textBox1.AppendText(Environment.NewLine);
                }
                string to = System.IO.Path.Combine(extractPath, fileName + ".kml");
                foreach (var kml in Directory.EnumerateFiles(tempPath, "*.kml"))
                {
                    try
                    {
                        System.IO.File.Move(kml, to);
                    }
                    catch (Exception ex)
                    {
                        textBox1.AppendText(ex.Message);
                        textBox1.AppendText(kml);
                        textBox1.AppendText(Environment.NewLine);
                    }
                }

                // clear all file and folder in Temp folder
                System.IO.DirectoryInfo di = new DirectoryInfo(tempPath);

                foreach (FileInfo fileToDelete in di.GetFiles())
                {
                    fileToDelete.Delete();
                }
                foreach (DirectoryInfo dir in di.GetDirectories())
                {
                    dir.Delete(true);
                }
            }

            XNamespace ns = "http://www.opengis.net/kml/2.2";
            foreach (var file in Directory.EnumerateFiles(extractPath, "*.kml"))
            {
                XDocument doc = XDocument.Load(file);
                var query = doc.Root
                   .Element(ns + "Document")
                   .Elements(ns + "Placemark")
                   .Select(x => new PlaceMark // I assume you've already got this
                   {
                       Name = x.Element(ns + "name") != null ? x.Element(ns + "name").Value : "",
                       Description = x.Element(ns + "description") != null ? x.Element(ns + "description").Value : "",
                       longitude = x.Elements(ns + "LookAt").Select(c => c.Element(ns + "longitude")).FirstOrDefault() != null ?
                       x.Elements(ns + "LookAt").Select(c => c.Element(ns + "longitude")).FirstOrDefault().Value : "",
                       latitude = x.Elements(ns + "LookAt").Select(c=>c.Element(ns + "latitude")).FirstOrDefault() != null ?
                       x.Elements(ns + "LookAt").Select(c => c.Element(ns + "latitude")).FirstOrDefault().Value : "",
                       // etc
                   }).ToList();
            }
        }
    }

    internal class PlaceMark
    {
        public string Name { get; set; }
        public string Description { get; set; }
        public string longitude { get; set; }
        public string latitude { get; set; }
    }
}
