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

        /// <summary>
        /// Création d'une zone de désenfumage
        /// </summary>
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
            string sLayerNameG;
            string sLayerNameH;
            string sLayerNameT;
            
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Layer table for read
                LayerTable acLyrTbl;
                acLyrTbl = acTrans.GetObject(acCurDb.LayerTableId,
                                   OpenMode.ForRead) as LayerTable;

                sLayerNameG = string.Format("0 {0} SI PRO P - ZD{1} G", nameSociety, pStrRes.StringResult);
                sLayerNameH = string.Format("0 {0} SI PRO P - ZD{1} H", nameSociety, pStrRes.StringResult);
                sLayerNameT = string.Format("0 {0} SI PRO P - ZD{1} T", nameSociety, pStrRes.StringResult);

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
            
            //Création d'une zone
            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("");
            PromptResult pKeyRes;
            pKeyOpts.Message = string.Format("\nCréation d'une polyligne pour la zone ZD{0} ", pStrRes.StringResult);
            pKeyOpts.Keywords.Add("Oui");
            pKeyOpts.Keywords.Add("Non");
            pKeyOpts.Keywords.Default = "Oui";
            pKeyOpts.AllowNone = false;

            pKeyRes = acDoc.Editor.GetKeywords(pKeyOpts);            
                        
            // Exit if the user presses ESC or cancels the command
            if (pKeyRes.Status == PromptStatus.Cancel)
                return;

            // Boucle tant que l'on veut créer une zone avec le même numéro
            bool createZone = true;
            while (createZone)
            {
                switch (pKeyRes.StringResult)
                {
                    case "Oui":
                        //Création d'une Polyligne sur le calque sLayerNameG
                        Polyline acPline = new Polyline();
                        acPline = Create_Pline(sLayerNameG);

                        //Création d'une hachure solid associé à la polyligne sur le calque sLayerNameH
                        //Démarre une transaction
                        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                        {
                            // Open the Block table for read
                            BlockTable acBlkTbl;
                            acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                            OpenMode.ForRead) as BlockTable;

                            // Open the Block table record Model space for write
                            BlockTableRecord acBlkTblRec;
                            acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                            OpenMode.ForWrite) as BlockTableRecord;

                            // Adds the pline to an object id array
                            ObjectIdCollection acObjIdColl = new ObjectIdCollection();
                            acObjIdColl.Add(acPline.ObjectId);

                            // Create the hatch object and append it to the block table record
                            using (Hatch acHatch = new Hatch())
                            {
                                acBlkTblRec.AppendEntity(acHatch);
                                acTrans.AddNewlyCreatedDBObject(acHatch, true);

                                // Set the properties of the hatch object
                                // Associative must be set after the hatch object is appended to the 
                                // block table record and before AppendLoop
                                acHatch.SetHatchPattern(HatchPatternType.PreDefined, "SOLID");
                                acHatch.Layer = sLayerNameH;
                                acHatch.Associative = true;
                                acHatch.AppendLoop(HatchLoopTypes.Outermost, acObjIdColl);
                                acHatch.EvaluateHatch(true);
                            }

                            // Save the changes and dispose of the transaction
                            acTrans.Commit();
                        }

                        //Demande si l'utilisateur veut recréer une zone
                        pKeyRes = acDoc.Editor.GetKeywords(pKeyOpts);
                        if (pKeyRes.StringResult == "Non")
                        {
                            createZone = false;
                            return;
                        }
                        break;

                    case "Non":
                        createZone = false;
                        return;
                }
            }

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

                // Boucle tant que l'utilisateur veut rajouter des lignes
                while (clickPoint)
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
                        default:
                            pPtOpts.BasePoint = lstPt[i - 1];
                            break;
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
        /// Création d'un Polyligne sur un calque défini
        /// </summary>
        /// <param name="layerName">Nom du claque</param>
        /// <returns></returns>
        static Polyline Create_Pline(string layerName)
        {            
            Document acDoc = AcAp.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;
            
            Matrix3d ucs = acDoc.Editor.CurrentUserCoordinateSystem;
            Plane plane = new Plane(Point3d.Origin, ucs.CoordinateSystem3d.Zaxis);

            Point3dCollection cPt = new Point3dCollection();
            Polyline acPline = new Polyline();

            PromptPointResult pPtRes;
            PromptPointOptions pPtOpts = new PromptPointOptions("");

            try
            {
                // Saisie du premier point
                pPtOpts.Message = "\nSpécifiez le premier point :";
                pPtRes = acDoc.Editor.GetPoint(pPtOpts);

                // Exit if the user presses ESC or cancels the command
                if (pPtRes.Status == PromptStatus.Cancel)
                {
                    return acPline;
                }          

                // Déplacement du SCU si le Z du point est différent de 0.0
                if (pPtRes.Value.Z != 0.0)
                {
                    Vector3d disp = new Vector3d(0.0, 0.0, pPtRes.Value.Z).TransformBy(ucs);
                    acDoc.Editor.CurrentUserCoordinateSystem = ucs.PreMultiplyBy(Matrix3d.Displacement(disp));
                }
                // Elevation de la polyligne
                double elev = pPtRes.Value
                    .TransformBy(ucs)
                    .TransformBy(Matrix3d.WorldToPlane(ucs.CoordinateSystem3d.Zaxis))
                    .Z;                

                //start a transaction
                using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                {
                    BlockTable acBlkTbl;
                    BlockTableRecord acBlkTblRec;

                    // Open Model space for write
                    acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId,
                                                 OpenMode.ForRead) as BlockTable;
                    acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace],
                                                    OpenMode.ForWrite) as BlockTableRecord;

                    // Define the new Pline
                    acPline.SetDatabaseDefaults();
                    acPline.Normal = ucs.CoordinateSystem3d.Zaxis;
                    acPline.Elevation = elev;
                    acPline.Layer = layerName;
                    acPline.AddVertexAt(0, pPtRes.Value.TransformBy(ucs).Convert2d(plane), 0.0, 0.0, 0.0);

                    // Add the line to the drawing
                    acBlkTblRec.AppendEntity(acPline);
                    acTrans.AddNewlyCreatedDBObject(acPline, true);

                    // Commit the changes and dispose of the transaction
                    acTrans.Commit();
                }
                //Ajoute le point dans la liste des points
                cPt.Add(pPtRes.Value);
                int i = 1;

                // Boucle tant que l'utilisateur veut cliquer des points
                bool clickPoint = true;
                while (clickPoint)
                {
                    switch (cPt.Count)
                    {
                        case 0:
                            pPtOpts.Message = "\nSpécifiez le premier point: ";
                            pPtOpts.Keywords.Clear();
                            pPtOpts.UseBasePoint = false;
                            break;
                        case 1:
                            pPtOpts.Message = "\nSpécifiez le point suivant ou ";
                            pPtOpts.Keywords.Clear();
                            pPtOpts.Keywords.Add("annUler");
                            pPtOpts.UseBasePoint = true;
                            pPtOpts.BasePoint = cPt[i - 1];
                            break;
                        case 2:
                            pPtOpts.Message = "\nSpécifiez le point suivant ou ";
                            pPtOpts.Keywords.Clear();
                            pPtOpts.Keywords.Add("annUler");
                            pPtOpts.BasePoint = cPt[i - 1];
                            break;
                        case 3:
                            pPtOpts.Message = "\nSpécifiez le point suivant ou ";
                            pPtOpts.Keywords.Clear();
                            pPtOpts.Keywords.Add("annUler");
                            pPtOpts.Keywords.Add("Clore");
                            pPtOpts.BasePoint = cPt[i - 1];
                            break;
                        default:
                            pPtOpts.BasePoint = cPt[i - 1];
                            break;
                    }
                    //Saisie du point suivant
                    pPtRes = acDoc.Editor.GetPoint(pPtOpts);
                    

                    switch (pPtRes.Status)
                    {
                        case PromptStatus.Cancel:
                            return acPline;
                        case PromptStatus.OK:
                            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                            {
                                acTrans.GetObject(acPline.ObjectId, OpenMode.ForWrite);
                                acPline.AddVertexAt(i, pPtRes.Value.TransformBy(ucs).Convert2d(plane), 0.0, 0.0, 0.0);
                                //Code à améliorer
                                if (i==0)
                                {
                                    acPline.RemoveVertexAt(1);
                                }
                                // Fin du code à améliorer
                                acTrans.Commit();
                            }
                            cPt.Add(pPtRes.Value);
                            i++;
                            break;
                        case PromptStatus.Keyword:
                            switch (pPtRes.StringResult)
                            {
                                case "Clore":
                                    using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                    {
                                        acTrans.GetObject(acPline.ObjectId, OpenMode.ForWrite);
                                        acPline.AddVertexAt(i, cPt[0].TransformBy(ucs).Convert2d(plane), 0.0, 0.0, 0.0);
                                        // Close the polyline
                                        acPline.Closed = true;
                                        acTrans.Commit();
                                    }                                    
                                    i++;
                                    clickPoint = false;
                                    break;
                                case "annUler":
                                    if (cPt.Count == 1)
                                    {
                                        cPt.RemoveAt(i - 1);
                                        i--;
                                    }
                                    else
                                    {
                                        using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
                                        {
                                            acTrans.GetObject(acPline.ObjectId, OpenMode.ForWrite);
                                            acPline.RemoveVertexAt(i - 1);
                                            acTrans.Commit();
                                        }
                                        cPt.RemoveAt(i - 1);
                                        i--;
                                        break;
                                    }                                    
                                    break;
                            }
                            break;
                    }
                }
                // Fin de la boucle
            }
            catch (System.Exception e)
            {
                acDoc.Editor.WriteMessage("\nErreur: {0}", e.Message);
            }
            finally
            {
                acDoc.Editor.CurrentUserCoordinateSystem = ucs;
            }
            return acPline;

        }

    }
}
