using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

[assembly: CommandClass(typeof(TestTransactionManager.TransactionManagerClass))]

namespace TestTransactionManager
{
    public class TransactionManagerClass : IExtensionApplication
    {
        public void Initialize()
        {
            //throw new NotImplementedException();
        }

        [CommandMethod("MAKETEST")]
        public void OpenTransactionManager()
        {
            // Get the current document and database
            Document acDoc = Application.DocumentManager.MdiActiveDocument;
            Database acCurDb = acDoc.Database;

            // Start a transaction
            using Transaction acTrans = acCurDb.TransactionManager.StartTransaction();
            // Open the Block table for read
            BlockTable? acBlkTbl = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead) as BlockTable;
            if (acBlkTbl != null)
            {
                // Open the Block table record Model space for read
                BlockTableRecord? acBlkTblRec = acTrans.GetObject(acBlkTbl[BlockTableRecord.ModelSpace], OpenMode.ForRead) as BlockTableRecord;
                if (acBlkTblRec != null)
                {
                    // Step through the Block table record
                    foreach (ObjectId asObjId in acBlkTblRec)
                    {
                        if (acTrans.GetObject(asObjId, OpenMode.ForRead) is Table tbl)
                        {
                            var tableName = tbl.Name;
                            var tableColumnsCount = tbl.Columns.Count;
                            var tableRowsCount = tbl.Rows.Count;
                            acDoc.Editor.WriteMessage("\nDXF name: " + asObjId.ObjectClass.DxfName);
                            acDoc.Editor.WriteMessage("\nObjectID: " + asObjId.ToString());
                            acDoc.Editor.WriteMessage("\nHandle: " + asObjId.Handle.ToString());
                            acDoc.Editor.WriteMessage("\nTableName: " + tbl.Name.ToString());
                            acDoc.Editor.WriteMessage("\nColumnsCount: " + tbl.Columns.Count.ToString());
                            var colWidths = "";
                            for (var j = 0; j < tbl.Columns.Count; j++)
                            {
                                colWidths += "[" + tbl.Columns[j].Width + "][" + tbl.Columns[j].Alignment + "]";
                            }
                            acDoc.Editor.WriteMessage("\nColumnWidthst: " + colWidths);
                            acDoc.Editor.WriteMessage("\nRowsCount: " + tbl.Rows.Count.ToString());
                            for (var i = 0; i < tbl.Rows.Count; i++)
                            {
                                for (var j = 0; j < tbl.Columns.Count; j++)
                                {
                                   var cell = tbl.Cells[i, j];
                                    acDoc.Editor.WriteMessage($"\nCells[{i}, {j}]: [{cell.Alignment}][{cell.TextHeight}]{cell.TextString}");
                                }
                            }
                            acDoc.Editor.WriteMessage("\n");
                        }
                    }
                    acTrans.Commit();
                }
            }
            // Dispose of the transaction
        }

        public void Terminate()
        {
            //throw new NotImplementedException();
        }
    }
}
