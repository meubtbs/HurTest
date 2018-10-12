using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace HurTest
{
    public partial class HurTestRibbon
    {
        private void HurTestRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnInsertGroup_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisDocument.SoruGrubuEkle();
        }

        private void btnInsertQuestion_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisDocument.SoruEkle();
        }

        private void btnInsertNumeric_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisDocument.InsertNumericField();
        }

        private void btnInsertChoice_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisDocument.InsertChoiceField();
        }
    }
}
