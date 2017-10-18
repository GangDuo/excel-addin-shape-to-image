using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace Excel.AddIn.Shape2Image
{
    public partial class Ribbon
    {
        private void Ribbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void RibbonButton_Click(object sender, RibbonControlEventArgs e)
        {
            RibbonButton control = sender as RibbonButton;
            new Liaison() { ControlId = control.Id }.Entrust();
        }
    }
}
