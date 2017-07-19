/* Copyright (C) 2017 Rosso Fabien

This program is free software; you can redistribute it and/or modify
it under the terms of the GNU General Public License as published by
the Free Software Foundation; either version 2 of the License, or
(at your option) any later version.

This program is distributed in the hope that it will be useful,
but WITHOUT ANY WARRANTY; without even the implied warranty of
MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the
GNU General Public License for more details.

You should have received a copy of the GNU General Public License along
with this program; if not, write to the Free Software Foundation, Inc.,
51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.*/


using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Colors;

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;


namespace SsiCad
{
    public class Commands
    {
        string nameSociety = "OTEIS";

        [CommandMethod("ZD")]
        public void Zone_Desenfumage()
        {
            Document acDoc = AcAp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            //Saisie utilisateur: récupère le numéro de la ZD
            PromptStringOptions pStrOpts = new PromptStringOptions("\nNuméro de la zone de désenfumage");
            pStrOpts.AllowSpaces = true;
            PromptResult pStrRes = acDoc.Editor.GetString(pStrOpts);

            // Exit if the user presses ESC or cancels the command
            if (pStrRes.Status == PromptStatus.Cancel)
                return;

            //Création des Claques pour la ZD si il n'existe pas
            //Démarre une transaction
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                   OpenMode.ForRead) as LayerTable;

                string sLayerNameG = string.Format("0 {0} SI PRO P - ZD{1} G", nameSociety, pStrRes.StringResult);
                string sLayerNameH = string.Format("0 {0} SI PRO P - ZD{1} H", nameSociety, pStrRes.StringResult);
                string sLayerNameT = string.Format("0 {0} SI PRO P - ZD{1} T", nameSociety, pStrRes.StringResult);

                List<string> lLayerName = new List<string>();
                lLayerName.Add(sLayerNameG);
                lLayerName.Add(sLayerNameH);
                lLayerName.Add(sLayerNameT);

                foreach (string sLayerName in lLayerName)
                {
                    if (acLyrTbl.Has(sLayerName) == false)
                    {
                        LayerTableRecord acLyrTblRec = new LayerTableRecord();

                        // Assign the layer the ACI color 1 and a name
                        acLyrTblRec.Color = Color.FromColorIndex(ColorMethod.ByAci, 1);
                        acLyrTblRec.Name = sLayerName;

                        // Upgrade the Layer table for write
                        acLyrTbl.UpgradeOpen();

                        // Append the new layer to the Layer table and the transaction
                        acLyrTbl.Add(acLyrTblRec);
                        acTrans.AddNewlyCreatedDBObject(acLyrTblRec, true);
                    }
                }

                // Save the changes and dispose of the transaction
                acTrans.Commit();
            }
            Create_Line();
            
            //Création d'une zone
            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
            pKeyOpts.Message = string.Format("\nCréation d'une polyligne pour la zone ZD{0} ", pStrRes.StringResult);
            pKeyOpts.Keywords.Add("Oui");
            pKeyOpts.Keywords.Add("Non");
            pKeyOpts.Keywords.Default = "Oui";
            pKeyOpts.AllowNone = false;

            PromptResult pKeyRes = acDoc.Editor.GetKeywords(pKeyOpts);

            // Exit if the user presses ESC or cancels the command
            if (pKeyRes.Status == PromptStatus.Cancel)
                return;            
        }

        /// <summary>
        /// Création d'une ou plusieurs lignes
        /// </summary>
        /// <returns>Retourne la liste des lignes créées</returns>
        static List<Line> Create_Line()
        {
            Document acDoc = AcAp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");
            List<Point3d> lstPt = new List<Point3d>();
            List<Line> lstAcLine = new List<Line>();

            // Saisie du premier point
            pPtOpts.Message = "\nSpécifiez le premier point: ";
            pPtRes = acDoc.Editor.GetPoint(pPtOpts);
            lstPt.Add(pPtRes.Value);

            // Exit if the user presses ESC or cancels the command
            if (pPtRes.Status == PromptStatus.Cancel)
            {
                return lstAcLine;
            }

            // Déplacement du SCU si le Z du point est différent de 0.0

            Matrix3d ucs = acDoc.Editor.CurrentUserCoordinateSystem;
            if (pPtRes.Value.Z != 0.0)
            {
                Vector3d disp = new Vector3d(0.0, 0.0, pPtRes.Value.Z).TransformBy(ucs);
                acDoc.Editor.CurrentUserCoordinateSystem = ucs.PreMultiplyBy(Matrix3d.Displacement(disp));
            }

            try
            {
                int i = 1;
                Boolean clickPoint = true;

                while (clickPoint == true)
                {
                    switch (lstPt.Count)
                    {
                        case 0:
                            pPtOpts.Message = "\nSpécifiez le premier point: ";
                            pPtOpts.UseBasePoint = false;
                            break;
                        case 1:
                            pPtOpts.Message = "\nSpécifiez le point suivant";
                            pPtOpts.Keywords.Clear();
                            pPtOpts.Keywords.Add("annUler");
                            pPtOpts.UseBasePoint = true;
                            pPtOpts.BasePoint = lstPt[i - 1];
                            break;
                        case 2:
                            pPtOpts.Keywords.Clear();
                            pPtOpts.Keywords.Add("annUler");
                            pPtOpts.BasePoint = lstPt[i - 1];
                            break;
                        case 3:
                            pPtOpts.Message = "\nSpécifiez le point suivant ou ";
                            pPtOpts.Keywords.Clear();
                            pPtOpts.Keywords.Add("annUler");
                            pPtOpts.Keywords.Add("Clore");
                            pPtOpts.BasePoint = lstPt[i - 1];
                            break;
                    }

                    if (lstPt.Count >= 4)
                    {
                        pPtOpts.BasePoint = lstPt[i - 1];
                    }

                    // Saisie du point suivant                    

                    pPtRes = acDoc.Editor.GetPoint(pPtOpts);

                    switch (pPtRes.Status)
                    {
                        case PromptStatus.OK:
                            if (lstPt.Count == 0)
                            {
                                lstPt.Add(pPtRes.Value);
                                i++;
                            }
                            else
                            {
                                // Start a transaction
                                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                {
                                    BlockTable acBlkTbl;
                                    BlockTableRecord acBlkTblRec;


                                    // Open Model space for write
                                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                                 OpenMode.ForRead) as BlockTable;
                                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                                    OpenMode.ForWrite) as BlockTableRecord;

                                    // Define the new line
                                    Line acLine = new Line();
                                    acLine.StartPoint = lstPt[i - 1];
                                    acLine.EndPoint = pPtRes.Value;

                                    acLine.SetDatabaseDefaults();
                                    lstAcLine.Add(acLine);

                                    // Add the line to the drawing
                                    acBlkTblRec.AppendEntity(acLine);
                                    acTrans.AddNewlyCreatedDBObject(acLine, true);

                                    // Commit the changes and dispose of the transaction
                                    acTrans.Commit();
                                }
                                lstPt.Add(pPtRes.Value);
                                i++;
                            }
                            break;
                        case PromptStatus.Keyword:
                            switch (pPtRes.StringResult)
                            {
                                case "Clore":
                                    // Start a transaction
                                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                    {
                                        BlockTable acBlkTbl;
                                        BlockTableRecord acBlkTblRec;

                                        // Open Model space for write
                                        acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                                     OpenMode.ForRead) as BlockTable;
                                        acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                                        OpenMode.ForWrite) as BlockTableRecord;

                                        // Define the new line
                                        Line acLine = new Line();
                                        acLine.StartPoint = lstPt[i - 1];
                                        acLine.EndPoint = lstPt[0];

                                        acLine.SetDatabaseDefaults();
                                        lstAcLine.Add(acLine);

                                        // Add the line to the drawing
                                        acBlkTblRec.AppendEntity(acLine);
                                        acTrans.AddNewlyCreatedDBObject(acLine, true);

                                        // Commit the changes and dispose of the transaction
                                        acTrans.Commit();
                                    }
                                    i++;
                                    clickPoint = false;

                                    break;
                                case "annUler":
                                    if (lstAcLine.Count == 0)
                                    {
                                        lstPt.RemoveAt(i - 1);
                                        i--;
                                    }
                                    else
                                    {
                                        ObjectId entId = lstAcLine[i - 2].ObjectId;
                                        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                        {
                                            //check if entity is already in erase state..
                                            if (entId.IsErased)
                                            {
                                                //GetObject, 3rd parameter openErased
                                                Entity ent = (Entity)acTrans.GetObject(entId,
                                                                               OpenMode.ForWrite, true);
                                                ent.Erase(false);
                                                entId = ObjectId.Null;
                                            }
                                            else
                                            {
                                                Entity ent = (Entity)acTrans.GetObject(entId,
                                                                                  OpenMode.ForWrite);
                                                ent.Erase();
                                            }
                                            acTrans.Commit();
                                        }
                                        lstPt.Add(pPtRes.Value);
                                        lstPt.RemoveAt(i - 1);
                                        lstAcLine.RemoveAt(i - 2);
                                        i--;
                                    }
                                    break;
                            }
                            break;
                        case PromptStatus.Cancel:
                            return lstAcLine;
                    }
                }
            }
            catch (System.Exception e)
            {
                acDoc.Editor.WriteMessage("\nErreur: {0}", e.Message);
            }
            finally
            {
                acDoc.Editor.CurrentUserCoordinateSystem = ucs;
            }
            return lstAcLine;
        }

        /// <summary>
        /// Création d'une Polyligne
        /// </summary>
        /// <returns>Retourne l'object polylignes créées</returns>
        static void Create_Pline()
        {
            Document acDoc = AcAp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Saisie du premier point
            PromptPointOptions ppo = new PromptPointOptions("\nPoint de départ: ");
            PromptPointResult ppr = acDoc.Editor.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK)
                return;

            // Déplacement du SCU si le Z du point est différent de 0.0
            Matrix3d ucs = acDoc.Editor.CurrentUserCoordinateSystem;
            Point3d p0 = ppr.Value;
            if (p0.Z != 0.0)
            {
                Vector3d disp = new Vector3d(0.0, 0.0, p0.Z).TransformBy(ucs);
                acDoc.Editor.CurrentUserCoordinateSystem = ucs.PreMultiplyBy(Matrix3d.Displacement(disp));
            }
            // Elevation de la polyligne
            double elev = p0
                .TransformBy(ucs)
                .TransformBy(Matrix3d.WorldToPlane(ucs.CoordinateSystem3d.Zaxis))
                .Z;

            // Plan de construction
            Plane plane = new Plane(Point3d.Origin, ucs.CoordinateSystem3d.Zaxis);
            // Saisie du second point
            ppo.Message = "\nPoint suivant: ";
            ppo.BasePoint = p0;
            ppo.UseBasePoint = true;
            ppo.AllowNone = true;
            ppr = acDoc.Editor.GetPoint(ppo);
            if (ppr.Status != PromptStatus.OK)
                return;

            Point3d pt = ppr.Value;
            try
            {
                // Création de la polyligne à 2 sommets
                ObjectId plId = ObjectId.Null;
                int i = 1;
                using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
                {
                    Polyline pline = new Polyline();
                    pline.Normal = ucs.CoordinateSystem3d.Zaxis;
                    pline.Elevation = elev;
                    pline.AddVertexAt(0, p0.TransformBy(ucs).Convert2d(plane), 0.0, 0.0, 0.0);
                    pline.AddVertexAt(1, pt.TransformBy(ucs).Convert2d(plane), 0.0, 0.0, 0.0);
                    BlockTableRecord btr =
                        (BlockTableRecord)tr.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite);
                    plId = btr.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);
                    tr.Commit();
                }
                // Ajout des points suivants
                while (true)
                {
                    ppo.BasePoint = pt;
                    ppr = acDoc.Editor.GetPoint(ppo);
                    if (ppr.Status != PromptStatus.OK)
                        break;

                    pt = ppr.Value;
                    i++;

                    using (Transaction tr = acCurDb.TransactionManager.StartTransaction())
                    {
                        Polyline pl = (Polyline)tr.GetObject(plId, OpenMode.ForWrite);
                        pl.AddVertexAt(i, pt.TransformBy(ucs).Convert2d(plane), 0.0, 0.0, 0.0);
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception e)
            {
                acDoc.Editor.WriteMessage("\nErreur: {0}", e.Message);
            }
            finally
            {
                acDoc.Editor.CurrentUserCoordinateSystem = ucs;
            }                       
        }
    }
}
