using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

namespace DocumentControlToolbar
{
    public partial class DocumentControlRibbon
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void docPropUpdater_Click(object sender, RibbonControlEventArgs e) {
            new DocPropertiesEditor().Show();
        }

        private void runAcronymTool_Click(object sender, RibbonControlEventArgs e) {
            new AcronymTableTool();
        }

        private void boilerplateFormat_Click(object sender, RibbonControlEventArgs e) {

        }
    }
}
