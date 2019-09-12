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
        public string sLayout = "";
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

        public void InsertBlockInAutoCAD(string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng, string units)
        {
            var blockPath = @"C:\ACAD2018_SMITH-NEPHEW\Standards\TitleBlock.dwg";
            var blockName = "MAINPATIENTDETAILS";
            if (!File.Exists(blockPath))
            {
                MessageBox.Show("Title Block Drawing does not exists");
            }

            //InsertDwg(blockPath, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);
            AddBlockToDB(blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units);
            //ImportBlocks(blockPath, blockName);
            //InsertImportedBlock(blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units);
            //InsertDwg(blockPath, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);
            //UpdateAttributes(doc.Database, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng);
        }

        private void InsertImportedBlock(string blockName, string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng, string units)
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
                BlockReference acBlkRef = default(BlockReference);
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

                            //using (AttributeDefinition acAttDef = new AttributeDefinition())
                            //{
                            //    acAttDef.Position = new Point3d(0, 0, 0);
                            //    acAttDef.Verifiable = false;
                            //    acAttDef.Prompt = ":";
                            //    acAttDef.Tag = "Door#";
                            //    acAttDef.TextString = "DXX";
                            //    acAttDef.Height = 1;
                            //    acAttDef
                            //    acAttDef.Justify = AttachmentPoint.MiddleCenter;

                            //    acBlkTblRec.AppendEntity(acAttDef);

                            //    acBlkTbl.UpgradeOpen();
                            //    acBlkTbl.Add(acBlkTblRec);
                            //    acTrans.AddNewlyCreatedDBObject(acBlkTblRec, true);
                            //}
                        }
                    }



                    // Save the new object to the database
                    acTrans.Commit();
                }
            }
            UpdateAttributes(doc.Database, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units);
        }


        public void AddBlockToDB(string blockName, string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng, string units)

        {

            Database db = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Database;
            String strLayout = BlockTableRecord.ModelSpace;
            if(sLayout.ToUpper().Equals("YES", StringComparison.InvariantCultureIgnoreCase))
            {
                strLayout = BlockTableRecord.PaperSpace;
            }
            using (Transaction myT = db.TransactionManager.StartTransaction())

            {

                //Get the block definition "Check".

                BlockTable bt = db.BlockTableId.GetObject(OpenMode.ForRead) as BlockTable;

                BlockTableRecord blockDef = bt[blockName].GetObject(OpenMode.ForRead) as BlockTableRecord;

                //Also open modelspace - we'll be adding our BlockReference to it

                //----------BlockTableRecord ms = bt[BlockTableRecord.ModelSpace].GetObject(OpenMode.ForWrite) as BlockTableRecord;
                BlockTableRecord ms = bt[strLayout].GetObject(OpenMode.ForWrite) as BlockTableRecord;

                //Create new BlockReference, and link it to our block definition

                Point3d point = new Point3d(2.0, 4.0, 6.0);

                using (BlockReference blockRef = new BlockReference(point, blockDef.ObjectId))
                {

                    //Add the block reference to modelspace
                    ms.AppendEntity(blockRef);

                    myT.AddNewlyCreatedDBObject(blockRef, true);

                    //Iterate block definition to find all non-constant

                    // AttributeDefinitions

                    foreach (ObjectId id in blockDef)

                    {

                        DBObject obj = id.GetObject(OpenMode.ForRead);

                        AttributeDefinition attDef = obj as AttributeDefinition;

                        if ((attDef != null) && (!attDef.Constant))

                        {

                            //This is a non-constant AttributeDefinition

                            //Create a new AttributeReference

                            using (AttributeReference attRef = new AttributeReference())

                            {

                                attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                                if (attRef.Tag.ToUpper().Equals("FILENO", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //attRef.UpgradeOpen();
                                    attRef.TextString = fileNo;
                                    //attRef.DowngradeOpen();
                                }
                                if (attRef.Tag.ToUpper().Equals("PATIENT", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //attRef.UpgradeOpen();
                                    attRef.TextString = patient;
                                    //attRef.DowngradeOpen();
                                }
                                if (attRef.Tag.ToUpper().Equals("DIAGNOSIS", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //attRef.UpgradeOpen();
                                    attRef.TextString = diagnosis;
                                    //attRef.DowngradeOpen();
                                }
                                if (attRef.Tag.ToUpper().Equals("AGE", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //attRef.UpgradeOpen();
                                    attRef.TextString = age;
                                    //attRef.DowngradeOpen();
                                }
                                if (attRef.Tag.ToUpper().Equals("SEX", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //attRef.UpgradeOpen();
                                    attRef.TextString = sex;
                                    //attRef.DowngradeOpen();
                                }
                                if (attRef.Tag.ToUpper().Equals("WORKORDER", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //attRef.UpgradeOpen();
                                    attRef.TextString = workOrder;
                                    //attRef.DowngradeOpen();
                                }
                                if (attRef.Tag.ToUpper().Equals("ORDERDATE", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    //attRef.UpgradeOpen();
                                    attRef.TextString = tempDate;
                                    //attRef.DowngradeOpen();
                                }
                                if (attRef.Tag.ToUpper().Equals("TEMPLATEENGINEER", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    // //attRef.UpgradeOpen();
                                    attRef.TextString = tempEng;
                                    // //attRef.DowngradeOpen();
                                }
                                if (attRef.Tag.ToUpper().Equals("UNITS", StringComparison.InvariantCultureIgnoreCase))
                                {
                                    // //attRef.UpgradeOpen();
                                    attRef.TextString = units;
                                    // //attRef.DowngradeOpen();
                                }
                                // attRef.TextString = "1";

                                //Add the AttributeReference to the BlockReference

                                blockRef.AttributeCollection.AppendAttribute(attRef);

                                myT.AddNewlyCreatedDBObject(attRef, true);

                            }

                        }

                    }

                }

                //Our work here is done

                myT.Commit();

            }

        }


        public void InsertDwg(string fname, string blockName, string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng, string units)
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

            UpdateAttributes(doc.Database, blockName, fileNo, patient, diagnosis, age, sex, workOrder, tempDate, tempEng, units);

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
                myT.Commit();
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


        private void UpdateAttributes(Database db, string blockName, string fileNo, string patient, string diagnosis, string age, string sex, string workOrder, string tempDate, string tempEng, string units)
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
                                                attRef.TextString = fileNo;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("PATIENT", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = patient;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("DIAGNOSIS", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = diagnosis;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("AGE", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = age;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("SEX", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = sex;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("WORKORDER", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = workOrder;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("ORDERDATE", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = tempDate;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("TEMPLATEENGINEER", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = tempEng;
                                                attRef.DowngradeOpen();
                                            }
                                            if (attRef.Tag.ToUpper().Equals("UNITS", StringComparison.InvariantCultureIgnoreCase))
                                            {
                                                attRef.UpgradeOpen();
                                                attRef.TextString = units;
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

        public void BindAttributes(ObjectId blkId, ref string fileNo, ref string patient, ref string diagnosis, ref string age, ref string sex, ref string workOrder, ref string tempDate, ref string tempEng, ref string units)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;
            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockReference blkRef = (BlockReference)tr.GetObject(blkId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blkRef.BlockTableRecord, OpenMode.ForRead);
                    //btr.Dispose();
                    Autodesk.AutoCAD.DatabaseServices.AttributeCollection attCol = blkRef.AttributeCollection;
                    int i = 1;
                    foreach (ObjectId attId in attCol)
                    {
                        // AttributeDetails _attr = new AttributeDetails();
                        AttributeReference attRef = (AttributeReference)tr.GetObject(attId, OpenMode.ForRead);
                        //attRef.SetAttributeFromBlock(attDef, blockRef.BlockTransform);
                        if (attRef.Tag.ToUpper().Equals("FILENO", StringComparison.InvariantCultureIgnoreCase))
                        {
                            fileNo = attRef.TextString;
                        }
                        if (attRef.Tag.ToUpper().Equals("PATIENT", StringComparison.InvariantCultureIgnoreCase))
                        {
                            patient = attRef.TextString;
                        }
                        if (attRef.Tag.ToUpper().Equals("DIAGNOSIS", StringComparison.InvariantCultureIgnoreCase))
                        {
                            diagnosis = attRef.TextString;
                        }
                        if (attRef.Tag.ToUpper().Equals("AGE", StringComparison.InvariantCultureIgnoreCase))
                        {
                            age = attRef.TextString;
                        }
                        if (attRef.Tag.ToUpper().Equals("SEX", StringComparison.InvariantCultureIgnoreCase))
                        {
                            sex = attRef.TextString;
                        }
                        if (attRef.Tag.ToUpper().Equals("WORKORDER", StringComparison.InvariantCultureIgnoreCase))
                        {
                            workOrder = attRef.TextString;
                        }
                        if (attRef.Tag.ToUpper().Equals("ORDERDATE", StringComparison.InvariantCultureIgnoreCase))
                        {
                            tempDate = attRef.TextString;
                        }
                        if (attRef.Tag.ToUpper().Equals("TEMPLATEENGINEER", StringComparison.InvariantCultureIgnoreCase))
                        {
                            tempEng = attRef.TextString;
                        }
                        if (attRef.Tag.ToUpper().Equals("UNITS", StringComparison.InvariantCultureIgnoreCase))
                        {
                            units = attRef.TextString;
                        }
                        i++;
                    }
                }
            }
            catch (System.Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        public ObjectId LoadBlockInstance(string BlockName = "MAINPATIENTDETAILS")
        {

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;


            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {

                    BlockTable BT = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    TypedValue[] typValue = new TypedValue[2];

                    typValue.SetValue(new TypedValue(Convert.ToInt32(DxfCode.Start), "INSERT"), 0);
                    typValue.SetValue(new TypedValue(Convert.ToInt32(DxfCode.BlockName), BlockName), 1);


                    PromptSelectionResult selRes = ed.SelectAll(new SelectionFilter(typValue));
                    if (selRes.Status != PromptStatus.OK)
                    {
                        ed.WriteMessage("\nNo " + BlockName + " block references selected");
                        return new ObjectId();
                    }
                    if (selRes.Value.Count != 0)
                    {
                        SelectionSet set = selRes.Value;
                        foreach (ObjectId id in set.GetObjectIds())
                        {
                            BlockReference oEnt = (BlockReference)tr.GetObject(id, OpenMode.ForRead);
                            return oEnt.ObjectId;
                            //Instance.Add(new InstancePosition { Position = "X :" + Math.Round(oEnt.Position.X, 2) + ", Y :" + Math.Round(oEnt.Position.Y, 2), blockObjId = oEnt.ObjectId });
                        }
                    }
                    tr.Commit();
                    return new ObjectId();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage(ex.ToString());
                return new ObjectId();
            }
        }
    }
}
