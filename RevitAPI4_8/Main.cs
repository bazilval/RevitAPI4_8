using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Mechanical;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPI4_8
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var arDoc = commandData.Application.ActiveUIDocument.Document;
            var ovDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ОВ")).FirstOrDefault();
            if (ovDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл ОВ");
                return Result.Failed;
            }
            var vkDoc = arDoc.Application.Documents.OfType<Document>().Where(x => x.Title.Contains("ВК")).FirstOrDefault();
            if (vkDoc == null)
            {
                TaskDialog.Show("Ошибка", "Не найден файл ВК");
                return Result.Failed;
            }
            View3D view3D = new FilteredElementCollector(arDoc)
                                                .OfClass(typeof(View3D))
                                                .OfType<View3D>()
                                                .Where(x => !x.IsTemplate)
                                                .FirstOrDefault();
            if (view3D == null)
            {
                TaskDialog.Show("Ошибка", "Не найден 3D вид");
                return Result.Failed;
            }

            FamilySymbol family = new FilteredElementCollector(arDoc)
                                    .OfClass(typeof(FamilySymbol))
                                    .OfCategory(BuiltInCategory.OST_GenericModel)
                                    .OfType<FamilySymbol>()
                                    .Where(x => x.FamilyName == "Отверстие")
                                    .FirstOrDefault();
            if (family == null)
            {
                TaskDialog.Show("Ошибка", "Не найдено семейство \"Отверстие\"");
                return Result.Failed;
            }

            List<Element> ducts = new FilteredElementCollector(ovDoc)
                                   .OfClass(typeof(Duct))
                                   .ToList();
            List<Element> pipes = new FilteredElementCollector(vkDoc)
                                   .OfClass(typeof(Pipe))
                                   .ToList();
            try
            {
                List<FamilyInstance> ovHoles = CreateHoles(arDoc, family, typeof(Duct), ducts, view3D);
                List<FamilyInstance> vkHoles = CreateHoles(arDoc, family, typeof(Pipe), pipes, view3D);
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }


            return Result.Succeeded;
        }

        private List<FamilyInstance> CreateHoles(Document arDoc, FamilySymbol family, Type type, List<Element> elements, View3D view3D)
        {
            ReferenceIntersector ri = new ReferenceIntersector(new ElementClassFilter(typeof(Wall)), FindReferenceTarget.Element, view3D);
            List<FamilyInstance> holes = new List<FamilyInstance>();

            if (!family.IsActive)
            {
                using (var ts = new Transaction(arDoc, "Activate of family"))
                {
                    ts.Start();
                    family.Activate();
                    ts.Commit();
                }
            }
            try
            {


                using (var ts = new Transaction(arDoc, "Создание отверстий"))
                {
                    ts.Start();


                    foreach (var elem in elements)
                    {
                        Line curve = (elem.Location as LocationCurve).Curve as Line;
                        XYZ point = curve.GetEndPoint(0);
                        XYZ direction = curve.Direction;

                        List<ReferenceWithContext> intersections = ri.Find(point, direction)
                                                    .Where(x => x.Proximity <= curve.Length)
                                                    .Distinct(new RWCEqualityCompaper())
                                                    .ToList();

                        foreach (var refer in intersections)
                        {
                            double proximity = refer.Proximity;
                            var reference = refer.GetReference();
                            Wall wall = arDoc.GetElement(reference.ElementId) as Wall;
                            Level level = arDoc.GetElement(wall.LevelId) as Level;
                            XYZ holePoint = point + (direction * proximity);

                            FamilyInstance hole = arDoc.Create.NewFamilyInstance(holePoint, family, wall, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
                            Parameter width = hole.LookupParameter("ADSK_Размер_Ширина");
                            Parameter height = hole.LookupParameter("ADSK_Размер_Высота");
                            if (type == typeof(Duct))
                            {
                                width.Set(GetDiam(elem as Duct));
                                height.Set(GetDiam(elem as Duct));
                            }
                            else if (type == typeof(Pipe))
                            {
                                width.Set(GetDiam(elem as Pipe));
                                height.Set(GetDiam(elem as Pipe));
                            }
                            holes.Add(hole);
                        }
                    }

                    ts.Commit();
                }
            }
            catch (Exception)
            {
                throw;
            }

            return holes;
        }

        public double GetDiam(Duct elem)
        {
            var duct = elem as Duct;
            return duct.Diameter;
        }

        public double GetDiam(Pipe elem)
        {
            var pipe = elem as Pipe;
            return pipe.Diameter;
        }
    }

    public class RWCEqualityCompaper : IEqualityComparer<ReferenceWithContext>
    {
        public bool Equals(ReferenceWithContext x, ReferenceWithContext y)
        {
            if (ReferenceEquals(x, y)) return true;
            if (ReferenceEquals(null, y)) return false;
            if (ReferenceEquals(null, x)) return false;

            var xRef = x.GetReference();
            var yRef = x.GetReference();
            return xRef.LinkedElementId == yRef.LinkedElementId &&
                    xRef.ElementId == yRef.ElementId;
        }

        public int GetHashCode(ReferenceWithContext obj)
        {
            var refer = obj.GetReference();

            unchecked
            {
                return (refer.LinkedElementId.GetHashCode() * 397) ^ refer.ElementId.GetHashCode();
            }
        }
    }
}
