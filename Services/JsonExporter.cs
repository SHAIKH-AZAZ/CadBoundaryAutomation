using System;
using System.IO;
using Autodesk.AutoCAD.DatabaseServices;
using Newtonsoft.Json;
using CadBoundaryAutomation.Models;

namespace CadBoundaryAutomation.Services
{
    public static class JsonExporter
    {
        public static string GetDefaultJsonPath(Database db)
        {
            string dwgPath = db.Filename;

            string folder;
            string name;

            if (!string.IsNullOrWhiteSpace(dwgPath))
            {
                folder = Path.GetDirectoryName(dwgPath);
                name = Path.GetFileNameWithoutExtension(dwgPath);
            }
            else
            {
                folder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                name = "Drawing";
            }

            return Path.Combine(folder, name + "_bars.json");
        }

        public static void Save(string path, BarsRunJson run)
        {
            string json = JsonConvert.SerializeObject(run, Formatting.Indented);
            File.WriteAllText(path, json);
        }
    }
}
