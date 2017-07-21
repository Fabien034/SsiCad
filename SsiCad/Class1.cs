using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;

namespace RegionalActivities
{
    public static class Extensions
    {
        // Region extensions

        ///<summary>
        /// Get the centroid of a Region.
        ///</summary>
        ///<param name="cur">An optional curve used to define the region.</param>
        ///<returns>A nullable Point3d containing the centroid of the Region.</returns>

        public static Point3d? GetCentroid(this Region reg, Curve cur = null)
        {
            if (cur == null)
            {
                var idc = new DBObjectCollection();
                reg.Explode(idc);
                if (idc.Count == 0)
                    return null;

                cur = idc[0] as Curve;
            }

            if (cur == null)
                return null;

            var cs = cur.GetPlane().GetCoordinateSystem();
            var o = cs.Origin;
            var x = cs.Xaxis;
            var y = cs.Yaxis;

            var a = reg.AreaProperties(ref o, ref x, ref y);
            var pl = new Plane(o, x, y);
            return pl.EvaluatePoint(a.Centroid);
        }
    }

    public class Commands
    {
        [CommandMethod("COR")]
        public void CentroidOfRegion()
        {
            var doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;
            var ed = doc.Editor;

            var peo = new PromptEntityOptions("\nSelect a region");
            peo.SetRejectMessage("\nMust be a region.");
            peo.AddAllowedClass(typeof(Region), false);
            var per = ed.GetEntity(peo);
            if (per.Status != PromptStatus.OK)
                return;

            using (var tr = doc.TransactionManager.StartTransaction())
            {
                var reg = tr.GetObject(per.ObjectId, OpenMode.ForRead) as Region;
                if (reg != null)
                {
                    var pt = reg.GetCentroid();
                    ed.WriteMessage("\nCentroid is {0}", pt);
                }
                tr.Commit();
            }
        }
    }
}