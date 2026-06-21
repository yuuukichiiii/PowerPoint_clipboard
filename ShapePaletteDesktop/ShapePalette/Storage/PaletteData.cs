using System;
using System.Collections.Generic;
using System.IO;
using Newtonsoft.Json;

namespace ShapePalette.Storage
{
    /// <summary>パレットの1アイテム。実体はライブラリ pptx のスライド（SlideID で参照）。</summary>
    public class PaletteItem
    {
        public string Id { get; set; } = Guid.NewGuid().ToString("N");
        public string Name { get; set; }
        public int SlideId { get; set; }       // ライブラリ pptx 内の安定 ID
        public string ThumbFile { get; set; }  // thumbs フォルダ内のファイル名
    }

    public class PaletteTab
    {
        public string Id { get; set; } = "t_" + Guid.NewGuid().ToString("N").Substring(0, 8);
        public string Name { get; set; }
        public List<PaletteItem> Items { get; set; } = new List<PaletteItem>();
    }

    public class PaletteData
    {
        public string ActiveTab { get; set; }
        public List<PaletteTab> Tabs { get; set; } = new List<PaletteTab>();

        public static string DataDir
        {
            get
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "ShapePalette");
                Directory.CreateDirectory(dir);
                return dir;
            }
        }

        public static string ThumbsDir
        {
            get { var d = Path.Combine(DataDir, "thumbs"); Directory.CreateDirectory(d); return d; }
        }

        public static string LibraryPath => Path.Combine(DataDir, "library.pptx");
        public static string DataPath => Path.Combine(DataDir, "palette.json");

        public PaletteTab ActiveTabObj()
        {
            PaletteTab t = null;
            if (!string.IsNullOrEmpty(ActiveTab))
                t = Tabs.Find(x => x.Id == ActiveTab);
            if (t == null && Tabs.Count > 0) { t = Tabs[0]; ActiveTab = t.Id; }
            return t;
        }

        public static PaletteData Load()
        {
            try
            {
                if (File.Exists(DataPath))
                {
                    var json = File.ReadAllText(DataPath);
                    var data = JsonConvert.DeserializeObject<PaletteData>(json);
                    if (data != null) { data.EnsureDefaults(); return data; }
                }
            }
            catch (Exception ex) { Log.Write("PaletteData.Load EX: " + ex.Message); }

            var fresh = new PaletteData();
            fresh.EnsureDefaults();
            return fresh;
        }

        public void Save()
        {
            try
            {
                var json = JsonConvert.SerializeObject(this, Formatting.Indented);
                File.WriteAllText(DataPath, json);
            }
            catch (Exception ex) { Log.Write("PaletteData.Save EX: " + ex.Message); }
        }

        private void EnsureDefaults()
        {
            if (Tabs == null) Tabs = new List<PaletteTab>();
            if (Tabs.Count == 0)
                Tabs.Add(new PaletteTab { Name = "図形" });
            if (string.IsNullOrEmpty(ActiveTab)) ActiveTab = Tabs[0].Id;
        }
    }
}
