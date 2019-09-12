using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;

namespace BlockCreation
{
    public class BlockCreation
    {
        private List<string> attributes = new List<string>();
        public BlockCreation()
        {
            attributes.Add("FILENO:");
            attributes.Add("PATIENT:");
            attributes.Add("DIAGNOSIS");
            attributes.Add("AGE:");
            attributes.Add("SEX:");
            attributes.Add("WORKORDER");
            attributes.Add("DATE");
            attributes.Add("TEMPENG");
            //attributes.Add("");
        }

        public void InsertBlockInAutoCAD(string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng)
        {
            var blockPath = @"C:\SN\TitleBlock.dwg";
            var blockName = "MAINPATIENTDETAILS";
            if (!File.Exists(blockPath))
            {
                MessageBox.Show("Title Block Drawing does not exists");
            }

            //InsertDwg(blockPath, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);

            ImportBlocks(blockPath, blockName);
            InsertImportedBlock(blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);
            //InsertDwg(blockPath, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);
            //UpdateAttributes(doc.Database, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);
        }

        private void InsertImportedBlock(string blockName, string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng)
        {
            Database acCurDb;
            acCurDb = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            using (Transaction acTrans = acCurDb.TransactionManager.StartTransaction())
            {
                // Open the Block table for read
                BlockTable acBlkTbl;
                acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;

                ObjectId blkRecId = ObjectId.Null;
                BlockReference acBlkRef=default(BlockReference);
                if (acBlkTbl.Has(blockName))
                {
                    blkRecId = acBlkTbl[blockName];
                    if (blkRecId != ObjectId.Null)
                    {
                       
                        using (acBlkRef = new BlockReference(new Point3d(0, 0, 0), blkRecId))
                        {
                            BlockTableRecord acCurSpaceBlkTblRec;
                            acCurSpaceBlkTblRec = acTrans.GetObject(acCurDb.CurrentSpaceId, OpenMode.ForWrite) as BlockTableRecord;

                            acCurSpaceBlkTblRec.AppendEntity(acBlkRef);
                            acTrans.AddNewlyCreatedDBObject(acBlkRef, true);
                        }
                    }
                    //var blockDefinition = (BlockTableRecord)acBlkTbl[blockName].GetObject(OpenMode.ForRead);
                    //foreach (var objectId in blockDefinition)
                    //{
                    //    var dbObject = objectId.GetObject(OpenMode.ForRead);
                    //    var attributeDefinition = dbObject as AttributeDefinition;
                    //    if ((attributeDefinition != null) && (!attributeDefinition.Constant))
                    //    {
                    //        if (attributes.Contains(attributeDefinition.Tag))
                    //        {
                    //            using (var attributeReference = new AttributeReference())
                    //            {
                    //                try
                    //                {
                    //                    attributeReference.SetAttributeFromBlock(attributeDefinition, acBlkRef.BlockTransform);
                    //                    attributeReference.Position = attributeDefinition.Position.TransformBy(acBlkRef.BlockTransform);
                    //                    //attributeReference.TextString = attributes.Select(attributeDefinition.Tag);


                    //                    acBlkRef.AttributeCollection.AppendAttribute(attributeReference);
                    //                    acTrans.AddNewlyCreatedDBObject(attributeReference, true);
                    //                }
                    //                catch (Autodesk.AutoCAD.Runtime.Exception exception)
                    //                {
                    //                    //Log.Logger.Error(exception, "Unable to set attribute {Attribute} on block {blockName}", attributeDefinition.TextString, blockName);
                    //                }
                    //            }
                    //        }
                    //    }
                    //}


                    // Save the new object to the database
                    acTrans.Commit();
                }
            }
            //UpdateAttributes(doc.Database, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);
        }


        public void InsertDwg(string fname, string blockName, string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // doc.GetLispSymbol("SF");
            //doc.GetLispSymbol("B:SYM-CIVIL");

            // string fname = "C:\\Users\\Jeff\\Documents\\Drawing1.dwg";
            ObjectId ObjId = default(ObjectId);
            using (doc.LockDocument())
            {
                using (Transaction trx = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)db.BlockTableId.GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead);
                    //BlockTableRecord btrMs = (BlockTableRecord)bt[BlockTableRecord.ModelSpace].GetObject(Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite);

                    BlockTableRecord btrMs = trx.GetObject(db.CurrentSpaceId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForWrite) as BlockTableRecord;


                    using (Database dbInsert = new Database(false, true))
                    {
                        dbInsert.ReadDwgFile(fname, System.IO.FileShare.Read, true, "");
                        ObjId = CheckBlockwithDrawingNameExists(fname, blockName);
                        if (ObjId.IsNull)
                            ObjId = db.Insert(Path.GetFileNameWithoutExtension(fname), dbInsert, true);
                    }
                    //PromptPointOptions ppo = new PromptPointOptions(Constants.vbCrLf + "Insertion Point");
                    //PromptPointResult ppr = default(PromptPointResult);
                    //ppr = ed.GetPoint(ppo);
                    //if (ppr.Status != PromptStatus.OK)
                    //{
                    //    ed.WriteMessage(Constants.vbCrLf + "You decided to QUIT!");
                    //    return;
                    //}
                    Point3d insertPt = new Point3d(5, 5, 0);
                    BlockReference bref = new BlockReference(insertPt, ObjId);
                    //bref.TransformBy(Matrix3d.Scaling(10, insertPt));
                    btrMs.AppendEntity(bref);
                    trx.AddNewlyCreatedDBObject(bref, true);



                    trx.Commit();
                }
            }

            UpdateAttributes(doc.Database, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);

        }

        public void ImportBlocks(string fileName, string blockName)
        {

            DocumentCollection dm = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager;

            Editor ed = dm.MdiActiveDocument.Editor;

            Database destDb = dm.MdiActiveDocument.Database;

            Database sourceDb = new Database(false, true);

            sourceDb.ReadDwgFile(fileName, System.IO.FileShare.Read, true, "");


            // Create a variable to store the list of block identifiers

            ObjectIdCollection blockIds = new ObjectIdCollection();


            Autodesk.AutoCAD.DatabaseServices.TransactionManager tm = sourceDb.TransactionManager;


            using (Transaction myT = tm.StartTransaction())

            {

                // Open the block table

                BlockTable bt = (BlockTable)tm.GetObject(sourceDb.BlockTableId, OpenMode.ForRead, false);


                // Check each block in the block table

                foreach (ObjectId btrId in bt)

                {

                    BlockTableRecord btr = (BlockTableRecord)tm.GetObject(btrId,
                                                    OpenMode.ForRead,
                                                    false);

                    // Only add named & non-layout blocks to the copy list

                    if (btr.Name.Equals(blockName, StringComparison.InvariantCultureIgnoreCase))

                        blockIds.Add(btrId);

                    btr.Dispose();

                }

            }

            // Copy blocks from source to destination database

            IdMapping mapping = new IdMapping();

            sourceDb.WblockCloneObjects(blockIds,

                                        destDb.BlockTableId,

                                        mapping,

                                        DuplicateRecordCloning.Replace,

                                        false);






            sourceDb.Dispose();

        }


        private void UpdateAttributes(Database db, string blockName, string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            ObjectId msId;//, psId;
            using (doc.LockDocument())
            {
                using (Transaction trans = db.TransactionManager.StartTransaction())
                {
                    BlockTable blockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead) as BlockTable;
                    msId = blockTable[BlockTableRecord.ModelSpace];

                    //psId = blockTable[BlockTableRecord.PaperSpace];
                    ObjectId objId = msId;// in blockTable)
                                          //{
                    BlockTableRecord blkTableRec = trans.GetObject(objId, OpenMode.ForRead) as BlockTableRecord;

                    foreach (ObjectId entId in blkTableRec)
                    {
                        Entity ent = trans.GetObject(entId, OpenMode.ForRead) as Entity;

                        if (ent != null)
                        {
                            BlockReference br = ent as BlockReference;
                            if (br != null)
                            {
                                //BlockTableRecord bd = (BlockTableRecord)trans.GetObject(br.BlockTableRecord, OpenMode.ForRead);
                                if (br.Name.ToUpper().Equals(blockName, StringComparison.InvariantCultureIgnoreCase))
                                {
                                    foreach (ObjectId arId in br.AttributeCollection)
                                    {
                                        DBObject obj = trans.GetObject(arId, OpenMode.ForRead);
                                        AttributeReference attRef = obj as AttributeReference;
                                        if (attRef != null)
                                        {
                                            if (attRef.Tag.ToUpper().Equals("FILENO", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = "File No:" + fileNo;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("PATIENT:", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = "Patient:" + patient;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("DIAGNOSIS:", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = "Diagnosis:" + diagnosis;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("AGE:", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = "Age:" + age;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("SEX:", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = "Sex:" + sex;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("WORKORDER:", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = "Work Order:" + workOrder;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("DATE:", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = "Date:" + tempDate;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("TEMPENG:", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = "Temp Eng:" + tempEng;
                                                attRef.DowngradeOpen();
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    trans.Commit();
                }
            }
        }

        private ObjectId CheckBlockwithDrawingNameExists(string path, string blockName)
        {
            ObjectId returnObjectId = default(ObjectId);
            Database sourceDb = new Database(false, true);
            sourceDb.ReadDwgFile(path, FileShare.Read, true, "");
            Autodesk.AutoCAD.DatabaseServices.TransactionManager transMan = sourceDb.TransactionManager;
            using (Transaction trans = transMan.StartTransaction())
            {
                BlockTable bt = (BlockTable)trans.GetObject(sourceDb.BlockTableId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, false);
                foreach (ObjectId btrId in bt)
                {
                    BlockTableRecord btr = (BlockTableRecord)trans.GetObject(btrId, Autodesk.AutoCAD.DatabaseServices.OpenMode.ForRead, false);
                    if (btr.Name.Equals(blockName))
                    {
                        returnObjectId = btrId;
                        trans.Commit();
                        break;
                    }
                }
            }
            return returnObjectId;
        }
    }
}
