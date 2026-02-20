using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointUsefulTools
{
    internal class CellStyleInfo
    {
        public Office.MsoFillType FillType { get; set; }
        public int FillForeColorRGB { get; set; }
        public float FillTransparency { get; set; }

        public string FontName { get; set; }
        public float FontSize { get; set; }
        public bool FontBold { get; set; }
        public bool FontItalic { get; set; }
        public int FontColorRGB { get; set; }

        public float MarginTop { get; set; }
        public float MarginBottom { get; set; }
        public float MarginLeft { get; set; }
        public float MarginRight { get; set; }

        public BorderInfo BorderTop { get; set; }
        public BorderInfo BorderBottom { get; set; }
        public BorderInfo BorderLeft { get; set; }
        public BorderInfo BorderRight { get; set; }
    }

    internal class BorderInfo
    {
        public bool Visible { get; set; }
        public int ColorRGB { get; set; }
        public float Weight { get; set; }
        public Office.MsoLineDashStyle DashStyle { get; set; }
    }

    internal class TableLayoutInfo
    {
        public CellStyleInfo HeaderStyle { get; set; }
        public CellStyleInfo BodyStyle { get; set; }
    }

    internal static class TableLayoutManager
    {
        public static void ApplyDefaultLayout(PowerPoint.Table table, DefaultTableSettings settings)
        {
            if (settings == null) return;
            var layout = new TableLayoutInfo
            {
                HeaderStyle = ToCellStyleInfo(settings.HeaderStyle),
                BodyStyle = ToCellStyleInfo(settings.BodyStyle)
            };
            ApplyLayout(table, layout);
        }

        private static CellStyleInfo ToCellStyleInfo(DefaultCellStyle style)
        {
            if (style == null) return null;
            return new CellStyleInfo
            {
                FillType = Office.MsoFillType.msoFillSolid,
                FillForeColorRGB = style.FillForeColorRGB,
                FillTransparency = style.FillTransparency,
                FontName = style.FontName,
                FontSize = style.FontSize,
                FontBold = style.FontBold,
                FontItalic = style.FontItalic,
                FontColorRGB = style.FontColorRGB,
                MarginTop = style.MarginTop,
                MarginBottom = style.MarginBottom,
                MarginLeft = style.MarginLeft,
                MarginRight = style.MarginRight,
                BorderTop = ToBorderInfo(style.BorderTop),
                BorderBottom = ToBorderInfo(style.BorderBottom),
                BorderLeft = ToBorderInfo(style.BorderLeft),
                BorderRight = ToBorderInfo(style.BorderRight)
            };
        }

        private static BorderInfo ToBorderInfo(DefaultBorderStyle style)
        {
            return new BorderInfo
            {
                Visible = style?.Visible ?? true,
                ColorRGB = style?.ColorRGB ?? 0x000000,
                Weight = style?.Weight ?? 0.75f,
                DashStyle = (Office.MsoLineDashStyle)(style?.DashStyle ?? 1)
            };
        }

        public static TableLayoutInfo CopyLayout(PowerPoint.Table table)
        {
            var info = new TableLayoutInfo();

            if (table.Rows.Count >= 1 && table.Columns.Count >= 1)
                info.HeaderStyle = CopyCellStyle(table.Cell(1, 1));

            if (table.Rows.Count >= 2 && table.Columns.Count >= 1)
                info.BodyStyle = CopyCellStyle(table.Cell(2, 1));
            else
                info.BodyStyle = info.HeaderStyle;

            return info;
        }

        private static CellStyleInfo CopyCellStyle(PowerPoint.Cell cell)
        {
            var style = new CellStyleInfo();

            var fill = cell.Shape.Fill;
            try { style.FillType = fill.Type; } catch { style.FillType = Office.MsoFillType.msoFillBackground; }
            try { style.FillForeColorRGB = fill.ForeColor.RGB; } catch { }
            try { style.FillTransparency = fill.Transparency; } catch { }

            var font = cell.Shape.TextFrame.TextRange.Font;
            try { style.FontName = font.Name; } catch { }
            try { style.FontSize = font.Size; } catch { }
            try { style.FontBold = font.Bold == Office.MsoTriState.msoTrue; } catch { }
            try { style.FontItalic = font.Italic == Office.MsoTriState.msoTrue; } catch { }
            try { style.FontColorRGB = font.Color.RGB; } catch { }

            var tf = cell.Shape.TextFrame;
            try { style.MarginTop = tf.MarginTop; } catch { }
            try { style.MarginBottom = tf.MarginBottom; } catch { }
            try { style.MarginLeft = tf.MarginLeft; } catch { }
            try { style.MarginRight = tf.MarginRight; } catch { }

            style.BorderTop = CopyBorder(cell, PowerPoint.PpBorderType.ppBorderTop);
            style.BorderBottom = CopyBorder(cell, PowerPoint.PpBorderType.ppBorderBottom);
            style.BorderLeft = CopyBorder(cell, PowerPoint.PpBorderType.ppBorderLeft);
            style.BorderRight = CopyBorder(cell, PowerPoint.PpBorderType.ppBorderRight);

            return style;
        }

        private static BorderInfo CopyBorder(PowerPoint.Cell cell, PowerPoint.PpBorderType borderType)
        {
            var info = new BorderInfo { Visible = true, Weight = 0.75f, DashStyle = Office.MsoLineDashStyle.msoLineSolid };
            try
            {
                dynamic border = cell.Borders[borderType];
                try { info.Visible = border.Visible != Office.MsoTriState.msoFalse; } catch { }
                try { info.ColorRGB = (int)border.ForeColor.RGB; } catch { }
                try { info.Weight = (float)border.Weight; } catch { }
                try { info.DashStyle = (Office.MsoLineDashStyle)(int)border.DashStyle; } catch { }
            }
            catch { }
            return info;
        }

        public static void ApplyLayout(PowerPoint.Table table, TableLayoutInfo layout)
        {
            if (layout == null) return;

            int rows = table.Rows.Count;
            int cols = table.Columns.Count;

            for (int r = 1; r <= rows; r++)
            {
                var style = (r == 1) ? layout.HeaderStyle : layout.BodyStyle;
                if (style == null) continue;

                for (int c = 1; c <= cols; c++)
                {
                    try { ApplyCellStyle(table.Cell(r, c), style); } catch { }
                }
            }
        }

        private static void ApplyCellStyle(PowerPoint.Cell cell, CellStyleInfo style)
        {
            var fill = cell.Shape.Fill;
            try
            {
                switch (style.FillType)
                {
                    case Office.MsoFillType.msoFillSolid:
                        fill.Solid();
                        fill.ForeColor.RGB = style.FillForeColorRGB;
                        fill.Transparency = style.FillTransparency;
                        break;
                    case Office.MsoFillType.msoFillBackground:
                        fill.Background();
                        break;
                    default:
                        fill.Solid();
                        fill.ForeColor.RGB = style.FillForeColorRGB;
                        fill.Transparency = style.FillTransparency;
                        break;
                }
            }
            catch { }

            var font = cell.Shape.TextFrame.TextRange.Font;
            try { if (!string.IsNullOrEmpty(style.FontName)) font.Name = style.FontName; } catch { }
            try { if (style.FontSize > 0) font.Size = style.FontSize; } catch { }
            try { font.Bold = style.FontBold ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; } catch { }
            try { font.Italic = style.FontItalic ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; } catch { }
            try { font.Color.RGB = style.FontColorRGB; } catch { }

            var tf = cell.Shape.TextFrame;
            try { tf.MarginTop = style.MarginTop; } catch { }
            try { tf.MarginBottom = style.MarginBottom; } catch { }
            try { tf.MarginLeft = style.MarginLeft; } catch { }
            try { tf.MarginRight = style.MarginRight; } catch { }

            ApplyBorder(cell, PowerPoint.PpBorderType.ppBorderTop, style.BorderTop);
            ApplyBorder(cell, PowerPoint.PpBorderType.ppBorderBottom, style.BorderBottom);
            ApplyBorder(cell, PowerPoint.PpBorderType.ppBorderLeft, style.BorderLeft);
            ApplyBorder(cell, PowerPoint.PpBorderType.ppBorderRight, style.BorderRight);
        }

        private static void ApplyBorder(PowerPoint.Cell cell, PowerPoint.PpBorderType borderType, BorderInfo info)
        {
            if (info == null) return;
            try
            {
                dynamic border = cell.Borders[borderType];
                try { border.Visible = info.Visible ? Office.MsoTriState.msoTrue : Office.MsoTriState.msoFalse; } catch { }
                if (info.Visible)
                {
                    try { border.ForeColor.RGB = info.ColorRGB; } catch { }
                    try { border.Weight = info.Weight; } catch { }
                    try { border.DashStyle = info.DashStyle; } catch { }
                }
            }
            catch { }
        }
    }
}
