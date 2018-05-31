using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace TestTaskVSTO
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            PowerPoint.Application thisApp = Globals.ThisAddIn.Application;
            PowerPoint.SlideRange slide = thisApp.ActiveWindow.Selection
                .SlideRange;
            var slideWidth = slide.CustomLayout.Width;
            
            foreach (PowerPoint.Shape shape in slide.Shapes)
            {
                if(shape.Type == Office.MsoShapeType.msoTextBox)
                {
                    if((shape.Width+shape.Left+10) <= slideWidth)
                    {
                        shape.IncrementLeft(10);
                    }
                    shape.TextEffect.FontBold = Office.MsoTriState.msoTrue;
                    shape.TextFrame.TextRange.Font.Color.RGB = 255;
                }
            }
        }
    }
}
