using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Excel.AddIn.Shape2Image
{
    class Liaison
    {
        public string ControlId { get; set; }

        public void Entrust()
        {
            var invokeAttr = BindingFlags.NonPublic | BindingFlags.InvokeMethod | BindingFlags.Static;
            this.GetType().InvokeMember(ControlId, invokeAttr, null, null, null);
        }

        private static void SaveAsPicture()
        {
            var workspace = new Workspace();
            workspace.Open();
            Worksheet activeSheet = Globals.ThisAddIn.Application.ActiveSheet;

            var shapes = activeSheet.Shapes.Cast<Shape>().Where(shape => new AllowedShape(shape).CanAllow()).ToArray();
            foreach (Shape shape in shapes)
            {
                // Shapeの右隣をファイル名として使用する
                var name = activeSheet.Cells[shape.TopLeftCell.Row, shape.TopLeftCell.Column + 1].Text;
                if (String.IsNullOrWhiteSpace(name)) continue;
                var file = Path.Combine(workspace.FullName, String.Format(@"{0}.jpg", name));
                var buf = new ShapeBuffer(shape);
                buf.SaveAsPicture(file);
            }
        }
    }
}
