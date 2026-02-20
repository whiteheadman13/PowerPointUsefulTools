using System;
using System.IO;
using System.Xml.Serialization;

namespace PowerPointUsefulTools
{
    [Serializable]
    public class DefaultBorderStyle
    {
        public bool Visible { get; set; }
        public int ColorRGB { get; set; }
        public float Weight { get; set; }
        public int DashStyle { get; set; } // MsoLineDashStyle: 1=実線, 2=点線, 3=破線, 4=一点鎖線, 5=二点鎖線
    }

    [Serializable]
    public class DefaultCellStyle
    {
        public int FillForeColorRGB { get; set; }
        public float FillTransparency { get; set; }
        public int FontColorRGB { get; set; }
        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
        public float MarginTop { get; set; }
        public float MarginBottom { get; set; }
        public float MarginLeft { get; set; }
        public float MarginRight { get; set; }

        public DefaultBorderStyle BorderTop { get; set; }
        public DefaultBorderStyle BorderBottom { get; set; }
        public DefaultBorderStyle BorderLeft { get; set; }
        public DefaultBorderStyle BorderRight { get; set; }
    }

    [Serializable]
    [XmlRoot("DefaultTableSettings")]
    public class DefaultTableSettings
    {
        public DefaultCellStyle HeaderStyle { get; set; }
        public DefaultCellStyle BodyStyle { get; set; }

        public DefaultTableSettings()
        {
            // Header: dark blue background (R=68, G=114, B=196), white text
            HeaderStyle = new DefaultCellStyle
            {
                FillForeColorRGB = 0xC47244,
                FillTransparency = 0f,
                FontColorRGB = 0xFFFFFF,
                FontName = "游ゴシック",
                FontSize = 11f,
                FontBold = false,
                FontItalic = false,
                MarginTop = 3.6f,
                MarginBottom = 3.6f,
                MarginLeft = 7.2f,
                MarginRight = 7.2f,
                BorderTop = new DefaultBorderStyle { Visible = true, ColorRGB = 0x000000, Weight = 0.75f, DashStyle = 1 },
                BorderBottom = new DefaultBorderStyle { Visible = true, ColorRGB = 0x000000, Weight = 0.75f, DashStyle = 1 },
                BorderLeft = new DefaultBorderStyle { Visible = true, ColorRGB = 0x000000, Weight = 0.75f, DashStyle = 1 },
                BorderRight = new DefaultBorderStyle { Visible = true, ColorRGB = 0x000000, Weight = 0.75f, DashStyle = 1 }
            };
            // Body: white background, black text
            BodyStyle = new DefaultCellStyle
            {
                FillForeColorRGB = 0xFFFFFF,
                FillTransparency = 0f,
                FontColorRGB = 0x000000,
                FontName = "游ゴシック",
                FontSize = 11f,
                FontBold = false,
                FontItalic = false,
                MarginTop = 3.6f,
                MarginBottom = 3.6f,
                MarginLeft = 7.2f,
                MarginRight = 7.2f,
                BorderTop = new DefaultBorderStyle { Visible = true, ColorRGB = 0x000000, Weight = 0.75f, DashStyle = 1 },
                BorderBottom = new DefaultBorderStyle { Visible = true, ColorRGB = 0x000000, Weight = 0.75f, DashStyle = 1 },
                BorderLeft = new DefaultBorderStyle { Visible = true, ColorRGB = 0x000000, Weight = 0.75f, DashStyle = 1 },
                BorderRight = new DefaultBorderStyle { Visible = true, ColorRGB = 0x000000, Weight = 0.75f, DashStyle = 1 }
            };
        }

        private static string SettingsFilePath => Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PowerPointUsefulTools",
            "DefaultTableSettings.xml");

        public static DefaultTableSettings Load()
        {
            var path = SettingsFilePath;
            if (!File.Exists(path))
                return new DefaultTableSettings();
            try
            {
                var serializer = new XmlSerializer(typeof(DefaultTableSettings));
                using (var reader = new StreamReader(path))
                    return (DefaultTableSettings)serializer.Deserialize(reader);
            }
            catch
            {
                return new DefaultTableSettings();
            }
        }

        public void Save()
        {
            var path = SettingsFilePath;
            Directory.CreateDirectory(Path.GetDirectoryName(path));
            var serializer = new XmlSerializer(typeof(DefaultTableSettings));
            using (var writer = new StreamWriter(path))
                serializer.Serialize(writer, this);
        }
    }
}
