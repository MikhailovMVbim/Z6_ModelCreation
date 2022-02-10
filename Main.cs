using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Z6_ModelCreation
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Получаем доступ к Revit, активному документу, базе данных документа
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;


            // запускаем транзакцию
            Transaction t = new Transaction(doc);
            t.Start("Create model");

            // получаем уровни
            Level levelBase = GetLevel(doc, "Уровень 1");
            Level levelTop = GetLevel(doc, "Уровень 2");
            Level levelRoof = GetLevel(doc, "Уровень 2");

            // создаем стены  с заданными габаритами на любом уровне и до любого уровня
            var walls = CreateWalls(doc, levelBase, levelTop, 9000, 6000);

            // создаем дверь в первой стене
            AddDoors(doc, levelBase, walls[0]);
            // создаем окна в оставшихся стенах
            for (int i = 1; i <= 3; i++)
            {
                AddWindows(doc, levelBase, walls[i]);
            }

            // создаем крышу по контуру
            //AddRoof(doc,levelRoof, walls);

            // создаем крышу выдавливанием
            AddExtrusionRoof(doc, levelRoof);

            t.Commit();

            return Result.Succeeded;
        }

        private void AddExtrusionRoof(Document doc, Level levelRoof)
        {
            RoofType roofType = new FilteredElementCollector(doc)
            .OfClass(typeof(RoofType))
            .OfType<RoofType>()
            .Where(d => d.Name.Equals("Типовой - 125мм"))
            .Where(f => f.FamilyName.Equals("Базовая крыша"))
            .FirstOrDefault();


            // создаем профиль выдавливания
            CurveArray profile = new CurveArray();
            profile.Append(Line.CreateBound(new XYZ(0, mmToFeet(-3500), mmToFeet(4000)), new XYZ(0, 0, mmToFeet(6000))));
            profile.Append(Line.CreateBound(new XYZ(0, 0, mmToFeet(6000)), new XYZ(0, mmToFeet(3500), mmToFeet(4000))));

            ReferencePlane refPlane = doc.Create.NewReferencePlane(new XYZ(0, 0, 0), new XYZ(0, 0, 20), new XYZ(0, 20, 0), doc.ActiveView);
            double extrusionStart = mmToFeet(-5000);
            double extrusionEnd = mmToFeet(5000);
            ExtrusionRoof extrusionRoof = doc.Create.NewExtrusionRoof(profile, refPlane, levelRoof, roofType, extrusionStart, extrusionEnd);
        }

        private void AddRoof(Document doc, Level levelRoof, List<Wall> walls)
        {
            RoofType roofType = new FilteredElementCollector(doc)
                .OfClass(typeof(RoofType))
                .OfType<RoofType>()
                .Where(d => d.Name.Equals("Типовой - 125мм"))
                .Where(f => f.FamilyName.Equals("Базовая крыша"))
                .FirstOrDefault();

            double widthWall = walls[0].Width;
            //double dt = widthWall * 0.5;
            double dt = mmToFeet(600); // добавим свес
            List<XYZ> points = new List<XYZ>(5);
            points.Add(new XYZ(-dt, -dt, 0.0));
            points.Add(new XYZ(-dt, dt, 0.0));
            points.Add(new XYZ(dt, dt, 0.0));
            points.Add(new XYZ(dt, -dt, 0.0));
            points.Add(points[0]);


            // получаем отпечаток для крыши со смещением
            CurveArray footPrint = doc.Application.Create.NewCurveArray();
            for (int i = 0; i < 4; i++)
            {
                LocationCurve locationCurve = walls[i].Location as LocationCurve;
                XYZ p1 = locationCurve.Curve.GetEndPoint(0) + points[i];
                XYZ p2 = locationCurve.Curve.GetEndPoint(1) + points[i + 1];
                Line line = Line.CreateBound(p1, p2);
                footPrint.Append(line);
            }
            ModelCurveArray mapping = new ModelCurveArray();
            FootPrintRoof footPrintRoof = doc.Create.NewFootPrintRoof(footPrint, levelRoof, roofType, out mapping);
            // делаем уклон
            foreach (ModelCurve modelCurve in mapping)
            {
                footPrintRoof.set_DefinesSlope(modelCurve, true);
                footPrintRoof.set_SlopeAngle(modelCurve, 0.5);
            }    
        }

        private static double mmToFeet(double lenght)
        {
            return UnitUtils.ConvertToInternalUnits(lenght, UnitTypeId.Millimeters);
        }

        private static List<Wall> CreateWalls(Document doc, Level levelBase, Level levelTop, double lenght, double width)
        {
            double dx = mmToFeet(lenght) * 0.5;
            double dy = mmToFeet(width) * 0.5;
            List<XYZ> wallPoints = new List<XYZ>();
            wallPoints.Add(new XYZ(-dx, -dy, 0));
            wallPoints.Add(new XYZ(-dx, dy, 0));
            wallPoints.Add(new XYZ(dx, dy, 0));
            wallPoints.Add(new XYZ(dx, -dy, 0));
            wallPoints.Add(wallPoints[0]);
            // список стен
            List<Wall> walls = new List<Wall>();
            // создаем стены на заданном уровне, высотой до требуемого уровня
            for (int i = 0; i < 4; i++)
            {
                Line line = Line.CreateBound(wallPoints[i], wallPoints[i + 1]);
                Wall wall = Wall.Create(doc, line, levelBase.Id, false);
                wall.get_Parameter(BuiltInParameter.WALL_HEIGHT_TYPE).Set(levelTop.Id);
                walls.Add(wall);
            }
            return walls;
        }

        private static void AddWindows(Document doc, Level levelBase, Wall wall)
        {
            FamilySymbol windowType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Windows)
                .OfType<FamilySymbol>()
                .Where(d => d.Name.Equals("0915 x 1830 мм"))
                .Where(f => f.FamilyName.Equals("Фиксированные"))
                .FirstOrDefault();
            if (!windowType.IsActive)
            {
                windowType.Activate();
            }
            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ insertPoint = (hostCurve.Curve.GetEndPoint(0) + hostCurve.Curve.GetEndPoint(1)) * 0.5;
            FamilyInstance window = doc.Create.NewFamilyInstance(insertPoint, windowType, wall, levelBase, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            window.get_Parameter(BuiltInParameter.INSTANCE_SILL_HEIGHT_PARAM).Set(mmToFeet(900));
        }

        private static void AddDoors(Document doc, Level levelBase, Wall wall)
        {
            FamilySymbol doorType = new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .OfCategory(BuiltInCategory.OST_Doors)
                .OfType<FamilySymbol>()
                .Where(d => d.Name.Equals("0915 x 2134 мм"))
                .Where(f => f.FamilyName.Equals("Одиночные-Щитовые"))
                .FirstOrDefault();
            if (!doorType.IsActive)
            {
                doorType.Activate();
            }
            LocationCurve hostCurve = wall.Location as LocationCurve;
            XYZ insertPoint = (hostCurve.Curve.GetEndPoint(0) + hostCurve.Curve.GetEndPoint(1)) * 0.5;
            doc.Create.NewFamilyInstance(insertPoint, doorType, wall, levelBase, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
        }

        private static Level GetLevel(Document doc, string levelName)
        {
            //получаем все уровни
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Level))
                .OfType<Level>()
                .Where(l => l.Name.Equals(levelName))
                .FirstOrDefault();
        }
    }
}
