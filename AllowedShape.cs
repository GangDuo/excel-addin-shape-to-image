using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Excel.AddIn.Shape2Image
{
    class AllowedShape
    {
        private readonly string[] AllowedShapeName = new string[] { "Picture", "Group", "図" };
        
        public Shape Shape { get; set; }

        public AllowedShape(Shape shape)
        {
            Shape = shape;
        }

        public bool CanAllow()
        {
            return AllowedShapeName.Any((x) => Shape.Name.StartsWith(x));
        }
    }
}
