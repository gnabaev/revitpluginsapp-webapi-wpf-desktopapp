﻿using AngleSharp.Html.Dom;
using AngleSharp.Html.Parser;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace RevitPluginsApp.Plugin.ClashManagement
{
    [Transaction(TransactionMode.Manual)]
    public class ClashIndicatorPlacementCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var fileNames = GetFileNamesFromDialog("Открыть HTML файл", "HTML Files (*.html)|*.html", false);

            if (fileNames.Length > 0)
            {
                Document doc = commandData.Application.ActiveUIDocument.Document;

                // Get coordinates of the document base point
                var projectPosition = GetDocumentBasePoint(doc);

                var projectPositionX = projectPosition.EastWest;
                var projectPositionY = projectPosition.NorthSouth;
                var projectPositionZ = projectPosition.Elevation;
                var projectPositionAngle = projectPosition.Angle;

                // Get the correct document title if the user is using a local copy of the central model
                var docTitle = GetDocumentTitle(doc);

                // Get the clash indicator family symbol
                var indicatorSymbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(f => f.FamilyName == "Индикатор коллизии");

                if (indicatorSymbol == null)
                {
                    TaskDialog.Show("Ошибка", $"В документе {docTitle} отсутствует семейство \"Индикатор коллизии\". Загрузите семейство и повторите попытку.");
                    return Result.Cancelled;
                }

                // Get the workset for clash indicators or create a new one
                var clashWorkset = GetWorkset(doc, "#Clashes");

                var errors = new List<string>();

                foreach (var fileName in fileNames)
                {
                    //Get HTML document tree
                    var htmlDoc = GetHtmlDocument(fileName);

                    var reportName = htmlDoc.QuerySelector(".testName").InnerHtml;

                    // Get the main table with clashes data
                    var mainTable = htmlDoc.QuerySelector(".mainTable");

                    var mainTableSection = mainTable.Children.FirstOrDefault();

                    // Get all rows of the main table
                    var mainTableRows = mainTableSection.Children;

                    var headerRows = mainTableRows.Where(i => i.ClassName == "headerRow").ToList();

                    var headerRow = headerRows[1];

                    var headerColumns = headerRow.Children;

                    // Get all general headers of the main table
                    var generalHeaderColumns = headerColumns.Where(c => c.ClassName == "generalHeader").ToList();

                    int clashIndex = -1;

                    int pointIndex = -1;

                    for (int i = 0; i < generalHeaderColumns.Count; i++)
                    {
                        if (generalHeaderColumns[i].InnerHtml == "Наименование конфликта")
                        {
                            clashIndex = i;
                        }

                        if (generalHeaderColumns[i].InnerHtml == "Точка конфликта")
                        {
                            pointIndex = i;
                        }
                    }

                    if (clashIndex == -1)
                    {
                        TaskDialog.Show("Ошибка", $"В главной таблице отчета {reportName} отсутствует столбец Наименование конфликта.");
                        return Result.Cancelled;
                    }

                    if (pointIndex == -1)
                    {
                        TaskDialog.Show("Ошибка", $"В главной таблице отчета {reportName} отсутствует столбец Точка конфликта.");
                        return Result.Cancelled;
                    }

                    // Get headers of the first element of the main table
                    var itemHeaderColumns = headerColumns.Where(c => c.ClassName == "item1Header").ToList();

                    int itemIdIndex = -1;

                    int itemModelIndex = -1;

                    for (int i = 0; i < itemHeaderColumns.Count; i++)
                    {
                        if (itemHeaderColumns[i].InnerHtml == "Id")
                        {
                            itemIdIndex = i;
                        }

                        if (itemHeaderColumns[i].InnerHtml == "Файл источника")
                        {
                            itemModelIndex = i;
                        }
                    }

                    if (itemIdIndex == -1)
                    {
                        TaskDialog.Show("Ошибка", $"В главной таблице отчета {reportName} отсутствует столбец Id.");
                        return Result.Cancelled;
                    }

                    if (itemModelIndex == -1)
                    {
                        TaskDialog.Show("Ошибка", $"В главной таблице отчета {reportName} отсутствует столбец Файл источника.");
                        return Result.Cancelled;
                    }

                    var contentRows = mainTableRows.Where(i => i.ClassName == "contentRow").ToList();

                    if (contentRows.Count == 0)
                    {
                        TaskDialog.Show("Предупреждение", $"В отчете {reportName} коллизий не обнаружено. Данный отчет будет пропущен.");
                        continue;
                    }

                    foreach (var contentRow in contentRows)
                    {
                        var contentColumns = contentRow.Children;

                        var clashName = contentColumns[clashIndex].InnerHtml;

                        var clashPointCoordinates = contentColumns[pointIndex].InnerHtml;

                        var clashPoint = GetClashPoint(clashPointCoordinates, projectPositionX, projectPositionY, projectPositionZ, projectPositionAngle);

                        var element1ContentColumns = contentColumns.Where(c => c.ClassName == "элемент1Содержимое").ToList();

                        var clashElementId1 = new ElementId(int.Parse(element1ContentColumns[itemIdIndex].InnerHtml));

                        var modelName1 = element1ContentColumns[itemModelIndex].InnerHtml;

                        var element2ContentColumns = contentColumns.Where(c => c.ClassName == "элемент2Содержимое").ToList();

                        var clashElementId2 = new ElementId(int.Parse(element2ContentColumns[itemIdIndex].InnerHtml));

                        var modelName2 = element2ContentColumns[itemModelIndex].InnerHtml;

                        if (modelName1 == docTitle && modelName2 == docTitle)
                        {
                            var element1 = doc.GetElement(clashElementId1);
                            var element2 = doc.GetElement(clashElementId2);

                            if (element1 != null && element2 != null)
                            {
                                using (Transaction transaction = new Transaction(doc))
                                {
                                    transaction.Start("Разместить индикатор коллизии");

                                    var indicatorInstance = PlaceClashIndicator(doc, clashPoint, indicatorSymbol, clashWorkset);

                                    FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                    transaction.Commit();
                                }
                            }
                            else
                            {
                                errors.Add($"{reportName} | {clashName} : Один или оба элемента не существуют в моделях");
                                continue;
                            }
                        }

                        else if (modelName1 != modelName2 && modelName1 == docTitle)
                        {
                            var rvtLinks = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

                            if (rvtLinks.Count == 0)
                            {
                                TaskDialog.Show("Ошибка", $"В текущем документе {docTitle} отсутствуют RVT-связи. Загрузите минимум одну RVT-связь для размещения индикаторов коллизий.");
                                return Result.Cancelled;
                            }

                            var linkDoc = rvtLinks.Select(l => l.GetLinkDocument()).Where(ld => ld != null).FirstOrDefault(ld => modelName2 == ld.Title + ".rvt");

                            if (linkDoc != null)
                            {
                                var element1 = doc.GetElement(clashElementId1);
                                var element2 = linkDoc.GetElement(clashElementId2);

                                if (element1 != null && element2 != null)
                                {
                                    using (Transaction transaction = new Transaction(doc))
                                    {
                                        transaction.Start("Разместить индикатор коллизии");

                                        var indicatorInstance = PlaceClashIndicator(doc, clashPoint, indicatorSymbol, clashWorkset);

                                        FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                        transaction.Commit();
                                    }
                                }
                                else
                                {
                                    errors.Add($"{reportName} | {clashName} : Один или оба элемента не существуют в моделях");
                                    continue;
                                }
                            }
                            else
                            {
                                errors.Add($"{reportName} | {clashName} : В текущем документе {docTitle} отсутствует или выгружена RVT-связь {modelName2}");
                                continue;
                            }
                        }

                        else if (modelName1 != modelName2 && modelName2 == docTitle)
                        {
                            var rvtLinks = new FilteredElementCollector(doc).OfClass(typeof(RevitLinkInstance)).Cast<RevitLinkInstance>().ToList();

                            if (rvtLinks.Count == 0)
                            {
                                TaskDialog.Show("Ошибка", $"В текущем документе {docTitle} отсутствуют RVT-связи. Загрузите минимум одну RVT-связь для размещения индикаторов коллизий.");
                                return Result.Cancelled;
                            }

                            var linkDoc = rvtLinks.Select(l => l.GetLinkDocument()).Where(ld => ld != null).FirstOrDefault(ld => modelName1 == ld.Title + ".rvt");

                            if (linkDoc != null)
                            {
                                var element1 = linkDoc.GetElement(clashElementId1);
                                var element2 = doc.GetElement(clashElementId2);

                                if (element1 != null && element2 != null)
                                {
                                    using (Transaction transaction = new Transaction(doc))
                                    {
                                        transaction.Start("Разместить индикатор коллизии");

                                        var indicatorInstance = PlaceClashIndicator(doc, clashPoint, indicatorSymbol, clashWorkset);

                                        FillClashIndicatorInfo(indicatorInstance, reportName, clashName, clashElementId1, modelName1, clashElementId2, modelName2);

                                        transaction.Commit();
                                    }
                                }
                                else
                                {
                                    errors.Add($"{reportName} | {clashName} : Один или оба элемента не существуют в моделях");
                                    continue;
                                }
                            }
                            else
                            {
                                errors.Add($"{reportName} | {clashName} : В текущем документе {docTitle} отсутствует или выгружена RVT-связь {modelName1}");
                                continue;
                            }
                        }
                        else
                        {
                            errors.Add($"{reportName} | {clashName} : Наименования {modelName1} и {modelName2} не соответствуют текущему документу {docTitle}");
                            continue;
                        }
                    }

                    var errorLogPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

                    var errorLogName = $"Clashlog_{reportName}_{DateTime.Now:yyyy.MM.dd_HH-mm-ss}.txt";

                    if (errors.Count > 0)
                    {
                        using (var errorLog = new StreamWriter(Path.Combine(errorLogPath, errorLogName)))
                        {
                            foreach (var error in errors)
                            {
                                errorLog.WriteLine(error);
                            }
                        }

                        TaskDialog.Show("Уведомление", $"Анализ отчета {reportName} завершен. Результаты анализа залогированы в файл {errorLogName} на рабочем столе.");
                    }
                    else
                    {
                        TaskDialog.Show("Уведомление", $"Анализ отчета {reportName} завершен. Все индикаторы размещены.");
                    }
                }
            }
            else
            {
                return Result.Cancelled;
            }

            return Result.Succeeded;
        }

        private string[] GetFileNamesFromDialog(string title, string filter, bool multiselect)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();

            openFileDialog.Title = title;
            openFileDialog.Filter = filter;
            openFileDialog.Multiselect = multiselect;

            bool? result = openFileDialog.ShowDialog();

            if (result == true)
            {
                return openFileDialog.FileNames;
            }
            else
            {
                return new string[0];
            }
        }

        private ProjectPosition GetDocumentBasePoint(Document doc)
        {
            var origin = new XYZ(0, 0, 0);

            var projectLocation = doc.ActiveProjectLocation;

            var projectPosition = projectLocation.GetProjectPosition(origin);

            return projectPosition;
        }

        private string GetDocumentTitle(Document doc)
        {
            string docTitle = doc.Title;

            if (!doc.IsWorkshared)
            {
                docTitle += ".rvt";
            }
            else
            {
                var separatorIndex = docTitle.LastIndexOf('_');

                if (separatorIndex != -1)
                {
                    docTitle = docTitle.Substring(0, separatorIndex) + ".rvt";
                }
            }

            return docTitle;
        }

        private Workset GetWorkset(Document doc, string worksetName)
        {
            Workset workset = null;

            if (doc.IsWorkshared)
            {
                var worksets = new FilteredWorksetCollector(doc).ToWorksets();

                var desiredWorkset = worksets.FirstOrDefault(x => x.Name == worksetName);

                if (desiredWorkset != null)
                {
                    workset = desiredWorkset;
                }
                else
                {
                    using (Transaction transaction = new Transaction(doc))
                    {
                        transaction.Start("Создать рабочий набор");

                        workset = Workset.Create(doc, worksetName);

                        transaction.Commit();
                    }
                }
            }

            return workset;
        }

        private IHtmlDocument GetHtmlDocument(string fileName)
        {
            string htmlFile = File.ReadAllText(fileName, Encoding.UTF8);

            IHtmlDocument htmlDoc = new HtmlParser().ParseDocument(htmlFile);

            return htmlDoc;
        }

        private XYZ GetClashPoint(string clashPointCoordinates, double projectPositionX, double projectPositionY, double projectPositionZ, double projectPositionAngle)
        {
            var pointCoordinates = clashPointCoordinates.Split(',').SelectMany(x => x.Trim().Split(':')).Where((x, i) => i % 2 != 0).Select(x => UnitUtils.Convert(double.Parse(x.Replace('.', ',')), UnitTypeId.Meters, UnitTypeId.Feet)).ToArray();

            var reportClashPoint = new XYZ(pointCoordinates[0], pointCoordinates[1], pointCoordinates[2]);

            var rotation = Transform.CreateRotationAtPoint(XYZ.BasisZ, projectPositionAngle * (-1), new XYZ(projectPositionX, projectPositionY, projectPositionZ));

            var rotatedClashPoint = rotation.OfPoint(reportClashPoint);

            var movedClashPoint = new XYZ(rotatedClashPoint.X - projectPositionX, rotatedClashPoint.Y - projectPositionY, rotatedClashPoint.Z - projectPositionZ);

            return movedClashPoint;
        }

        private FamilyInstance PlaceClashIndicator(Document doc, XYZ point, FamilySymbol familySymbol, Workset workset)
        {
            if (!familySymbol.IsActive)
            {
                familySymbol.Activate();
            }

            var familyInstance = doc.Create.NewFamilyInstance(point, familySymbol, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);
            familyInstance.Pinned = true;
            familyInstance.get_Parameter(BuiltInParameter.ELEM_PARTITION_PARAM).Set(workset.Id.IntegerValue);

            return familyInstance;
        }

        private void FillClashIndicatorInfo(FamilyInstance clashIndicator, string reportName, string clashName, ElementId clashElementId1, string modelName1, ElementId clashElementId2, string modelName2)
        {
            clashIndicator.LookupParameter("Общ_Имя отчета").Set(reportName);
            clashIndicator.LookupParameter("Общ_Имя коллизии").Set(clashName);

            clashIndicator.LookupParameter("Общ_Идентификатор 1").Set(clashElementId1.ToString());
            clashIndicator.LookupParameter("Общ_Имя модели 1").Set(modelName1);

            clashIndicator.LookupParameter("Общ_Идентификатор 2").Set(clashElementId2.ToString());
            clashIndicator.LookupParameter("Общ_Имя модели 2").Set(modelName2);
        }
    }
}
