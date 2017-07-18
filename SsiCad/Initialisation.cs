using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace AutoCAD_2016_Plugin
{
    public class Initialization : IExtensionApplication
    {
        public void Initialize()
        {
            AcAp.Idle += OnIdle;
        }

        private void OnIdle(object sender, EventArgs e)
        {
            var doc = AcAp.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                AcAp.Idle -= OnIdle;
                doc.Editor.WriteMessage("\nSsiCad loaded.\n");
            }
        }

        public void Terminate()
        { }
    }
}
