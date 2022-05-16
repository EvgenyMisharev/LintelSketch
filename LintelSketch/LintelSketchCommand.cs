using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LintelSketch
{
    [Autodesk.Revit.Attributes.Transaction(Autodesk.Revit.Attributes.TransactionMode.Manual)]
    class LintelSketchCommand : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            Selection sel = commandData.Application.ActiveUIDocument.Selection;

            string assemblyPath = typeof(LintelSketchCommand).Assembly.Location;
            string assemblyFolder = Path.GetDirectoryName(assemblyPath);
            string libraryPath = Path.Combine(assemblyFolder, "data", "library");

            List<ElementId> lintelSelectionIds = sel.GetElementIds().ToList();

            View activeView = commandData.Application.ActiveUIDocument.ActiveView;
            ViewSchedule vs = activeView as ViewSchedule;
            if (vs == null)
            {
                message = "Перед запуском перейдите в Ведомость перемычек, выберите все строчки и после этого запустите плагин.";
                Debug.WriteLine("Active view is not ViewSchedule");
                return Result.Failed;
            }

            if (lintelSelectionIds.Count == 0)
            {
                message = "Перед запуском перейдите в Ведомость перемычек, выберите все строчки и после этого запустите плагин.";
                Debug.WriteLine("No selected elements");
                return Result.Failed;
            }

            List<FamilyInstance> lintelsList = GetLintelsFromCurrentSelection(doc, lintelSelectionIds);
            if (lintelsList.Count == 0)
            {
                message = "Не выбрано ни одного элемента. В свойствах Типа перемычек в параметре Группа модели укажите значение - Перемычки составные";
                Debug.WriteLine("Empty lintels list");
                return Result.Failed;
            }

            //Удаление ранее созданных картинок для данной ведомости деталей
            string imagesPrefix = vs.Id.IntegerValue.ToString();
            List<ElementId> oldImageIds = new FilteredElementCollector(doc)
                .WhereElementIsElementType()
                .OfClass(typeof(ImageType))
                .Where(i => i.Name.StartsWith(imagesPrefix))
                .Select(i => i.Id)
                .ToList();
            Debug.WriteLine("Old scetch images found: " + oldImageIds.Count.ToString());

            using (TransactionGroup tg = new TransactionGroup(doc))
            {
                tg.Start("Ведомость перемычек");

                if (oldImageIds.Count > 0)
                {
                    using (Transaction t = new Transaction(doc))
                    {
                        t.Start("Очистка");
                        doc.Delete(oldImageIds);
                        t.Commit();
                    }
                }

                using (Transaction t = new Transaction(doc))
                {
                    t.Start("Заполнение картинок");
                    //Добавление картинок перемычкам
                    foreach (FamilyInstance lint in lintelsList)
                    {
                        string scetchImageName = $"{imagesPrefix}_{lint.Symbol.Name}_{lint.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsValueString()}.bmp";
                        ImageType lintelImageType = new FilteredElementCollector(doc)
                            .WhereElementIsElementType()
                            .OfClass(typeof(ImageType))
                            .Cast<ImageType>()
                            .FirstOrDefault(i => i.Name == scetchImageName);
                        if(lintelImageType != null)
                        {
                            lint.LookupParameter("LintelImage").Set(lintelImageType.Id);
                        }
                        else
                        {
                            string scetchTemplatePath = Path.Combine(libraryPath, lint.Symbol.Family.Name, "scetch.png");
                            if(File.Exists(scetchTemplatePath))
                            {
                                string scetchTemplateFolderPath = Path.GetDirectoryName(scetchTemplatePath);
                                Bitmap templateImage = new Bitmap(scetchTemplatePath);
                                PixelFormat pixformat = templateImage.PixelFormat;
                                if (Enum.GetName(typeof(PixelFormat), pixformat).Contains("ndexed"))
                                {
                                    string msg = "INCORRECT IMAGE FORMAT: " + scetchTemplatePath.Replace("\\", " \\")
                                        + ", PLEASE RESAVE IMAGE WITH 24bit ColorDepth, PaintNET strongly recommended";
                                    Debug.WriteLine(msg);
                                    throw new Exception(msg);
                                }

                                Graphics gr = Graphics.FromImage(templateImage);
                                gr.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
                                gr.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Default;
                                gr.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.Default;

                                StringFormat format = StringFormat.GenericTypographic;
                                format.Alignment = StringAlignment.Center;
                                format.LineAlignment = StringAlignment.Center;

                                //Файл параметров
                                string parametersPath = Path.Combine(libraryPath, lint.Symbol.Family.Name, "parameters.txt");
                                List<string> textLines = File.ReadAllLines(parametersPath).ToList();
                                List<ElementId> processedElementsIdList = new List<ElementId>();
                                foreach (string str in textLines)
                                {
                                    if(textLines.IndexOf(str) != -1 && textLines.IndexOf(str) != 0 )
                                    {
                                        List<string > parametersList = str.Split(',').ToList();

                                        string familyOrParameterName = parametersList[0];
                                        float.TryParse(parametersList[1], out float b);
                                        float.TryParse(parametersList[2], out float h);
                                        float.TryParse(parametersList[3], out float angle);

                                        gr.TranslateTransform(b, h);
                                        gr.RotateTransform(-angle);

                                        Font fnt = new Font("Isocpeur", 45.0F, FontStyle.Regular, GraphicsUnit.Pixel);

                                        List<ElementId> depElementsIds = lint.GetDependentElements(new ElementCategoryFilter(BuiltInCategory.OST_GenericModel)).ToList();
                                        ElementId depElementId = depElementsIds.FirstOrDefault(de => (doc.GetElement(de) as FamilyInstance).Symbol.Family.Name == familyOrParameterName
                                        && !processedElementsIdList.Contains(de));

                                        if(depElementId != null)
                                        {
                                            gr.DrawString((doc.GetElement(depElementId) as FamilyInstance).get_Parameter(BuiltInParameter.ALL_MODEL_MARK).AsString(), fnt, Brushes.Black, 0, 0, format);
                                            processedElementsIdList.Add(depElementId);
                                            gr.RotateTransform(angle);
                                            gr.TranslateTransform(-b, -h);
                                        }

                                        else
                                        {
                                            Parameter lintInstanceParam = lint.LookupParameter(familyOrParameterName);
                                            if(lintInstanceParam != null)
                                            {
                                                if (((int)(lintInstanceParam.Definition as InternalDefinition).BuiltInParameter).Equals((int)BuiltInParameter.INSTANCE_ELEVATION_PARAM)
                                                    || ((int)(lintInstanceParam.Definition as InternalDefinition).BuiltInParameter).Equals((int)BuiltInParameter.INSTANCE_FREE_HOST_OFFSET_PARAM))
                                                {
                                                    string levelResult = "";
                                                    if (lintInstanceParam.AsDouble() > 0)
                                                    {
                                                        levelResult = "+" + Math.Round((lintInstanceParam.AsDouble() * 304.8 / 1000), 3).ToString("0.000");
                                                    }
                                                    else
                                                    {
                                                        levelResult = Math.Round((lintInstanceParam.AsDouble() * 304.8 / 1000), 3).ToString("0.000");
                                                    }

                                                    gr.DrawString(levelResult, fnt, Brushes.Black, 0, 0, format);
                                                    gr.RotateTransform(angle);
                                                    gr.TranslateTransform(-b, -h);
                                                    continue;
                                                }
                                                else
                                                {
                                                    gr.DrawString(lintInstanceParam.AsValueString(), fnt, Brushes.Black, 0, 0, format);
                                                    gr.RotateTransform(angle);
                                                    gr.TranslateTransform(-b, -h);
                                                    continue;
                                                }
                                            }

                                            lintInstanceParam = lint.Symbol.LookupParameter(familyOrParameterName);
                                            if (lintInstanceParam != null)
                                            {
                                                gr.DrawString(lintInstanceParam.AsValueString(), fnt, Brushes.Black, 0, 0, format);
                                                gr.RotateTransform(angle);
                                                gr.TranslateTransform(-b, -h);
                                            }
                                        }
                                    }
                                }
                                string symbolNameFormat = lint.Symbol.Name.Replace('/', '_');
                                string scetchImagePath = Path.Combine(scetchTemplateFolderPath, $"{imagesPrefix}_{symbolNameFormat}_{lint.get_Parameter(BuiltInParameter.INSTANCE_ELEVATION_PARAM).AsValueString()}.bmp");
                                templateImage.Save(scetchImagePath);

                                ImageType newLintelImageType = ImageType.Create(doc, scetchImagePath);
                                lint.LookupParameter("LintelImage").Set(newLintelImageType.Id);
                                File.Delete(scetchImagePath);
                            }
                        }
                    }

                    t.Commit();
                }
                tg.Assimilate();
            }
            return Result.Succeeded;
        }

        private static List<FamilyInstance> GetLintelsFromCurrentSelection(Document doc, List<ElementId> selIds)
        {
            List<FamilyInstance> tempLintelsList = new List<FamilyInstance>();
            foreach (ElementId lintelId in selIds)
            {
                if (doc.GetElement(lintelId) is FamilyInstance
                    && null != doc.GetElement(lintelId).Category
                    && doc.GetElement(lintelId).Category.Id.IntegerValue.Equals((int)BuiltInCategory.OST_GenericModel)
                    && ((doc.GetElement(lintelId) as FamilyInstance).Symbol.get_Parameter(BuiltInParameter.ALL_MODEL_MODEL).AsString() == "Перемычки составные"))
                {
                    tempLintelsList.Add(doc.GetElement(lintelId) as FamilyInstance);
                }
            }
            return tempLintelsList;
        }
    }
}
