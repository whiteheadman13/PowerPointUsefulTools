using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointUsefulTools
{
    public partial class ThisAddIn
    {
        internal TableLayoutInfo CopiedTableLayout { get; set; }
        internal int CopiedFromShapeId { get; set; } = -1;

        private bool _applyingLayout;

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            Application.WindowSelectionChange += Application_WindowSelectionChange;
        }

        private void Application_WindowSelectionChange(PowerPoint.Selection sel)
        {
            if (CopiedTableLayout == null || _applyingLayout) return;
            if (sel.Type != PowerPoint.PpSelectionType.ppSelectionShapes) return;

            PowerPoint.Shape tableShape = null;
            try
            {
                var shapeRange = sel.ShapeRange;
                if (shapeRange.Count > 0)
                {
                    var shape = shapeRange[1];
                    if (shape.HasTable == Office.MsoTriState.msoTrue)
                        tableShape = shape;
                }
            }
            catch { }

            if (tableShape == null || tableShape.Id == CopiedFromShapeId) return;

            _applyingLayout = true;
            try
            {
                TableLayoutManager.ApplyLayout(tableShape.Table, CopiedTableLayout);
            }
            finally
            {
                _applyingLayout = false;
            }
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonUsefulTools();
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
