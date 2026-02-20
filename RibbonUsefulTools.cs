using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointUsefulTools
{
    [ComVisible(true)]
    public class RibbonUsefulTools : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointUsefulTools.RibbonUsefulTools.xml");
        }

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        public void BtnTableLayoutCopy_Click(Office.IRibbonControl control)
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            var tableShape = GetSelectedTableShape(sel);
            if (tableShape != null)
            {
                Globals.ThisAddIn.CopiedTableLayout = TableLayoutManager.CopyLayout(tableShape.Table);
                Globals.ThisAddIn.CopiedFromShapeId = tableShape.Id;
            }
            else
            {
                Globals.ThisAddIn.CopiedTableLayout = null;
                Globals.ThisAddIn.CopiedFromShapeId = -1;
            }
        }

        private static PowerPoint.Shape GetSelectedTableShape(PowerPoint.Selection sel)
        {
            try
            {
                if (sel.Type == PowerPoint.PpSelectionType.ppSelectionShapes ||
                    sel.Type == PowerPoint.PpSelectionType.ppSelectionText)
                {
                    var shapeRange = sel.ShapeRange;
                    if (shapeRange.Count > 0)
                    {
                        var shape = shapeRange[1];
                        if (shape.HasTable == Office.MsoTriState.msoTrue)
                            return shape;
                    }
                }
            }
            catch { }
            return null;
        }

        public void BtnApplyDefaultLayout_Click(Office.IRibbonControl control)
        {
            var sel = Globals.ThisAddIn.Application.ActiveWindow.Selection;
            var tableShape = GetSelectedTableShape(sel);
            if (tableShape == null) return;

            var settings = DefaultTableSettings.Load();
            TableLayoutManager.ApplyDefaultLayout(tableShape.Table, settings);
        }

        public void BtnDefaultTableSettings_Click(Office.IRibbonControl control)
        {
            var settings = DefaultTableSettings.Load();
            using (var form = new TableSettingsForm(settings))
            {
                if (form.ShowDialog() == DialogResult.OK)
                    form.Result.Save();
            }
        }

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
