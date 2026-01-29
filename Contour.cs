using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

#pragma warning disable 0618

[assembly: CommandClass(typeof(Rivilis.Contour))]


namespace Rivilis
{
    public class Contour
    {
        [CommandMethod("MakePline")]
        public void MakePline()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;
            // Отбираем только отрезки и дуги
            SelectionFilter sf = new([new TypedValue((int)DxfCode.Start, "LINE,ARC")]);
            PromptSelectionResult rs = ed.GetSelection(sf);
            if (rs.Status != PromptStatus.OK) return;
            ObjectIdCollection ids = [.. rs.Value.GetObjectIds()];

            using Transaction tr = db.TransactionManager.StartTransaction();

            while (ids.Count > 0)
            {
                using Polyline? p = MakeJoinedPoly(tr, ref ids);
                if (p != null)
                {
                    using BlockTableRecord btr = (BlockTableRecord)tr.GetObject(db.CurrentSpaceId, OpenMode.ForWrite);
                    btr.AppendEntity(p);
                }
                else
                {
                    ed.WriteMessage("\nОшибка во входных данных!");
                    return;
                }
            }
            tr.Commit();
        }

        public static Polyline? MakeJoinedPoly(Transaction tr, ref ObjectIdCollection ids)
        {
            if (ids.Count == 0) return null;
            // Создаём полилинию
            Polyline p = new();
            p.SetDatabaseDefaults();
            ObjectId idFirst = ids[0];
            // Добавляем первые две вершины из первого выбранного примитива
            using Curve c = (Curve)tr.GetObject(idFirst, OpenMode.ForRead, false, true);
            p.AddVertexAt(0, ToPoint2d(c.StartPoint), BulgeFromArc(c, false), 0, 0);
            p.AddVertexAt(1, ToPoint2d(c.EndPoint), 0, 0, 0);
            Point3d nextPt = c.EndPoint;
            Point3d prevPt = c.StartPoint;
            ObjectId idOwn = c.OwnerId;

            ids.Remove(idFirst);

            int prevCnt = ids.Count + 1;

            while (ids.Count > 0 && ids.Count < prevCnt)
            {
                prevCnt = ids.Count;
                foreach (ObjectId id in ids)
                {
                    using Curve cv = (Curve)tr.GetObject(id, OpenMode.ForRead, false, true);
                    if (cv.StartPoint == nextPt || cv.EndPoint == nextPt)
                    {
                        double bulge = BulgeFromArc(cv, cv.EndPoint == nextPt);
                        p.SetBulgeAt(p.NumberOfVertices - 1, bulge);
                        if (cv.StartPoint == nextPt)
                            nextPt = cv.EndPoint;
                        else
                            nextPt = cv.StartPoint;
                        p.AddVertexAt(p.NumberOfVertices, ToPoint2d(nextPt), 0, 0, 0);
                        ids.Remove(id);
                        break;
                    }
                    else if (cv.StartPoint == prevPt || cv.EndPoint == prevPt)
                    {
                        double bulge = BulgeFromArc(cv, cv.StartPoint == prevPt);
                        if (cv.StartPoint == prevPt)
                            prevPt = cv.EndPoint;
                        else
                            prevPt = cv.StartPoint;
                        p.AddVertexAt(0, ToPoint2d(prevPt), bulge, 0, 0);
                        ids.Remove(id);
                        break;
                    }
                }
            }
            if (p.NumberOfVertices == 0)
                return null;
            else
                return p;
        }

        // Функция возвращает кривизну дуги (bulge) или 0.0
        static double BulgeFromArc(Curve c, bool clockwise)
        {
            double bulge = 0.0;
            Arc? a = c as Arc;
            if (a == null) return bulge;
            double newStart = (a.StartAngle > a.EndAngle) ?
              (a.StartAngle - 8 * Math.Atan(1)) : (a.StartAngle);
            bulge = Math.Tan((a.EndAngle - newStart) / 4);
            if (clockwise) bulge = -bulge;
            return bulge;
        }

        // Функция преобразует Point3d в Point2d (отбрасывает Z)
        static Point2d ToPoint2d(Point3d p)
        {
            return new Point2d(p.X, p.Y);
        }
    }
}
