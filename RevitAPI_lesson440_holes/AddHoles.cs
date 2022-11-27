using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.DB.Structure;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI_lesson440_holes
{

    [TransactionAttribute(TransactionMode.Manual)]

    public class AddHoles : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document arDoc = commandData.Application.ActiveUIDocument.Document;
            Document ovDoc = arDoc.Application
                                  .Documents
                                  .OfType<Document>()
                                  .Where(x=>x.Title.Contains("ОВ"))
                                  .FirstOrDefault();

            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден ОВ Файл");
                return Result.Cancelled;
            }

            FamilySymbol familySymbolHoles = new FilteredElementCollector(arDoc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_GenericModel)
                .OfType<FamilySymbol>()
                .Where(x => x.FamilyName.Equals("Отверстия"))
                .FirstOrDefault();

            if (familySymbolHoles == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстия\"");
                return Result.Cancelled;
            }

            List<Duct> ducts = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Duct))
                .OfType<Duct>()
                .ToList();

            List<Pipe> pipes = new FilteredElementCollector(ovDoc)
                .OfClass(typeof(Pipe))
                .OfType<Pipe>()
                .ToList();

            View3D view3D = new FilteredElementCollector(arDoc)
                                .OfClass(typeof(View3D))
                                .OfType<View3D>()
                                .Where(x => !x.IsTemplate)
                                .FirstOrDefault();

            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено 3D вид");
                return Result.Cancelled;
            }

            ReferenceIntersector referenceIntersector = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            
            using (Transaction tr = new Transaction(arDoc, "Активация семейства"))
            {
                tr.Start();

                if (!familySymbolHoles.IsActive)
                    familySymbolHoles.Activate();

                tr.Commit();
            }    

            using (Transaction transaction = new Transaction(arDoc, "Расстановка отверстий"))
            {
                transaction.Start();

                AddHolesByDuct(arDoc, familySymbolHoles, ducts, referenceIntersector);
                AddHolesByPipe(arDoc, familySymbolHoles, pipes, referenceIntersector);

                transaction.Commit();
            }

            return Result.Succeeded;
        }

        private static void AddHolesByDuct(Document Doc, FamilySymbol familySymbolHoles, List<Duct> ducts, ReferenceIntersector referenceIntersector)
        {
            foreach (Duct duct in ducts)
            {
                Line curve = (duct.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = (Doc.GetElement(reference.ElementId)) as Wall;
                    Level levelInsertHole = Doc.GetElement(wall.LevelId) as Level;

                    XYZ pointInsertHole = point + (direction * proximity);

                    FamilyInstance hole = Doc.Create.NewFamilyInstance(pointInsertHole, familySymbolHoles, wall, levelInsertHole, StructuralType.NonStructural);

                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(duct.Diameter);
                    height.Set(duct.Diameter);
                }
            }
        }

        private static void AddHolesByPipe(Document Doc, FamilySymbol familySymbolHoles, List<Pipe> pipes, ReferenceIntersector referenceIntersector)
        {
            foreach (Pipe pipe in pipes)
            {
                Line curve = (pipe.Location as LocationCurve).Curve as Line;
                XYZ point = curve.GetEndPoint(0);
                XYZ direction = curve.Direction;

                List<ReferenceWithContext> intersections = referenceIntersector.Find(point, direction)
                    .Where(x => x.Proximity <= curve.Length)
                    .Distinct(new ReferenceWithContextElementEqualityComparer())
                    .ToList();

                foreach (ReferenceWithContext refer in intersections)
                {
                    double proximity = refer.Proximity;
                    Reference reference = refer.GetReference();
                    Wall wall = (Doc.GetElement(reference.ElementId)) as Wall;
                    Level levelInsertHole = Doc.GetElement(wall.LevelId) as Level;

                    XYZ pointInsertHole = point + (direction * proximity);

                    FamilyInstance hole = Doc.Create.NewFamilyInstance(pointInsertHole, familySymbolHoles, wall, levelInsertHole, StructuralType.NonStructural);

                    Parameter width = hole.LookupParameter("Ширина");
                    Parameter height = hole.LookupParameter("Высота");
                    width.Set(pipe.Diameter);
                    height.Set(pipe.Diameter);
                }
            }
        }

        public class ReferenceWithContextElementEqualityComparer : IEqualityComparer<ReferenceWithContext>
        {
            public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
            {
                if (ReferenceEquals(x, y)) return true;
                if (ReferenceEquals(null, x)) return false;
                if (ReferenceEquals(null, y)) return false;

                var xReference = x.GetReference();

                var yReference = y.GetReference();

                return xReference.LinkedElementId == yReference.LinkedElementId
                           && xReference.ElementId == yReference.ElementId;
            }

            public int GetHashCode(ReferenceWithContext obj)
            {
                var reference = obj.GetReference();

                unchecked
                {
                    return (reference.LinkedElementId.GetHashCode() * 397) ^ reference.ElementId.GetHashCode();
                }
            }
        }

    }
}
