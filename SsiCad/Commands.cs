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

using AcAp = Autodesk.AutoCAD.ApplicationServices.Application;

namespace SsiCad
{
    public class Commands
    {        
        [CommandMethod("Cline")]
        public void Create_Line()
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
                return;
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
                    switch(lstPt.Count)
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
                            pPtOpts.BasePoint = lstPt[i-1];
                            break;
                        case 2:
                            pPtOpts.Keywords.Clear();
                            pPtOpts.Keywords.Add("annUler");
                            pPtOpts.BasePoint = lstPt[i-1];                            
                            break;
                        case 3:
                            pPtOpts.Message = "\nSpécifiez le point suivant ou ";
                            pPtOpts.Keywords.Clear();
                            pPtOpts.Keywords.Add("annUler");
                            pPtOpts.Keywords.Add("Clore");
                            pPtOpts.BasePoint = lstPt[i-1];
                            break;
                    }                    

                    if (lstPt.Count >= 4)
                    {
                        pPtOpts.BasePoint = lstPt[i-1];
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
                                        acLine.StartPoint = lstPt[i-1];
                                        acLine.EndPoint = lstPt[0];

                                        acLine.SetDatabaseDefaults();

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
                                        lstPt.RemoveAt(i-1);
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
                                        lstPt.RemoveAt(i - 1);
                                        lstAcLine.RemoveAt(i - 2);
                                        i--;
                                    }
                                    break;
                            }
                            break;
                        case PromptStatus.Cancel:
                            return;
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
