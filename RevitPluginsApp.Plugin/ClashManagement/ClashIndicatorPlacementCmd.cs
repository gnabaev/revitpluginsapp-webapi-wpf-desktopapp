using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace RevitPluginsApp.Plugin.ClashManagement
{
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ClashIndicatorPlacementCmd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            var fileNames = GetFileNamesFromDialog("Открыть HTML файл", "HTML Files (*.html)|*.html", false);

            // Checks if no report is selected
            if (fileNames.Length == 0) return Result.Cancelled;

            var doc = commandData.Application.ActiveUIDocument.Document;

            // Get coordinates of the document base point
            var projectPosition = GetDocumentBasePoint(doc);

            // Get the correct document title if the user is using a local copy of the central model
            var docTitle = GetDocumentTitle(doc);

            // Get the clash indicator family symbol
            var indicatorSymbol = GetClashIndicatorSymbol(doc);

            if (indicatorSymbol == null)
            {
                TaskDialog.Show("Ошибка", $"В документе {docTitle} отсутствует семейство \"Индикатор коллизии\". Загрузите семейство и повторите попытку.");
                return Result.Cancelled;
            }

            // Activate the indicator family symbol
            if (!indicatorSymbol.IsActive)
            {
                using (Transaction transaction = new Transaction(doc))
                {
                    transaction.Start("Активировать типоразмер индикатора");

                    indicatorSymbol.Activate();

                    transaction.Commit();
                }
            }

            // Get the workset for clash indicators or create a new one
            var clashWorkset = GetOrCreateWorkset(doc, "#Clashes");

            var rvtLinks = GetRevitLinks(doc);

            if (rvtLinks.Count == 0)
            {
                TaskDialog.Show("Ошибка", $"В текущем документе {docTitle} отсутствуют RVT-связи. Загрузите минимум одну RVT-связь для размещения индикаторов коллизий.");
                return Result.Cancelled;
            }

            // Form a path of a for error logging
            var errorLogPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            var errors = new List<string>();

            var htmlParser = new HtmlReportParser();

            foreach (var fileName in fileNames)
            {
                // Get Html document tree
                var htmlDoc = htmlParser.GetHtmlDocument(fileName);

                var reportName = htmlParser.GetReportName(htmlDoc);

                // Form a log filename by a report
                var errorLogName = $"Clashlog_{reportName}_{DateTime.Now:yyyy.MM.dd_HH-mm-ss}.txt";

                // Get the main table with clashes data
                var mainTable = htmlParser.GetReportMainTable(htmlDoc);

                // Get all rows of the main table
                var mainTableRows = htmlParser.GetMainTableRows(mainTable);

                // Get all headers of the main table
                var headerColumns = htmlParser.GetMainTableHeaderColumns(mainTableRows);

                // Get all general headers of the main table
                var generalHeaderColumns = htmlParser.GetGeneralHeaderColumns(headerColumns);

                int clashIndex = htmlParser.GetColumnIndex(generalHeaderColumns, "Наименование конфликта");

                if (clashIndex == -1)
                {
                    TaskDialog.Show("Ошибка", $"В главной таблице отчета {reportName} отсутствует столбец Наименование конфликта. Для решения проблемы обратитесь к BIM-координатору проекта.");
                    return Result.Cancelled;
                }

                int pointIndex = htmlParser.GetColumnIndex(generalHeaderColumns, "Точка конфликта");

                if (pointIndex == -1)
                {
                    TaskDialog.Show("Ошибка", $"В главной таблице отчета {reportName} отсутствует столбец Точка конфликта. Для решения проблемы обратитесь к BIM-координатору проекта.");
                    return Result.Cancelled;
                }

                // Get headers of the first element of the main table
                var itemHeaderColumns = htmlParser.GetItemHeaderColumns(headerColumns);

                int itemIdIndex = htmlParser.GetColumnIndex(itemHeaderColumns, "Id");

                if (itemIdIndex == -1)
                {
                    TaskDialog.Show("Ошибка", $"В главной таблице отчета {reportName} отсутствует столбец Id. Для решения проблемы обратитесь к BIM-координатору проекта.");
                    return Result.Cancelled;
                }

                int itemModelIndex = htmlParser.GetColumnIndex(itemHeaderColumns, "Файл источника");

                if (itemModelIndex == -1)
                {
                    TaskDialog.Show("Ошибка", $"В главной таблице отчета {reportName} отсутствует столбец Файл источника. Для решения проблемы обратитесь к BIM-координатору проекта.");
                    return Result.Cancelled;
                }

                // Get all table rows with clashes
                var contentRows = htmlParser.GetTableContentRows(mainTableRows);

                // Return if a report is empty
                if (contentRows.Count == 0)
                {
                    return Result.Cancelled;
                }

                foreach (var contentRow in contentRows)
                {
                    var contentColumns = contentRow.Children;

                    var clashName = contentColumns[clashIndex].InnerHtml;

                    var clashPointCoordinates = contentColumns[pointIndex].InnerHtml;

                    var clashPoint = ConvertClashPoint(clashPointCoordinates, projectPosition);

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
                        var linkDoc = rvtLinks.Select(l => l.GetLinkDocument())?.FirstOrDefault(ld => modelName2 == ld.Title + ".rvt");

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
                        var linkDoc = rvtLinks.Select(l => l.GetLinkDocument())?.FirstOrDefault(ld => modelName1 == ld.Title + ".rvt");

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

            return Result.Succeeded;
        }

        private string[] GetFileNamesFromDialog(string title, string filter, bool multiselect)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Title = title,
                Filter = filter,
                Multiselect = multiselect
            };

            return openFileDialog.ShowDialog() == true ? openFileDialog.FileNames : new string[0];
        }

        private ProjectPosition GetDocumentBasePoint(Document doc)
        {
            var projectLocation = doc.ActiveProjectLocation;

            return projectLocation.GetProjectPosition(XYZ.Zero);
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

        private FamilySymbol GetClashIndicatorSymbol(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>()
                .FirstOrDefault(f => f.FamilyName == "Индикатор коллизии");
        }

        private Workset GetOrCreateWorkset(Document doc, string worksetName)
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

        private List<RevitLinkInstance> GetRevitLinks(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(RevitLinkInstance))
                .Cast<RevitLinkInstance>()
                .ToList();
        }

        private XYZ ConvertClashPoint(string clashPointCoordinates, ProjectPosition basePoint)
        {
            var convertedCoordinates = clashPointCoordinates.Split(',').SelectMany(x => x.Trim().Split(':')).Where((x, i) => i % 2 != 0).Select(x => UnitUtils.Convert(double.Parse(x.Replace('.', ',')), UnitTypeId.Meters, UnitTypeId.Feet)).ToArray();

            var clashPoint = new XYZ(convertedCoordinates[0], convertedCoordinates[1], convertedCoordinates[2]);

            var rotation = Transform.CreateRotationAtPoint(XYZ.BasisZ, -basePoint.Angle, new XYZ(basePoint.EastWest, basePoint.NorthSouth, basePoint.Elevation));

            var rotatedClashPoint = rotation.OfPoint(clashPoint);

            var movedClashPoint = new XYZ(rotatedClashPoint.X - basePoint.EastWest, rotatedClashPoint.Y - basePoint.NorthSouth, rotatedClashPoint.Z - basePoint.Elevation);

            return movedClashPoint;
        }

        private FamilyInstance PlaceClashIndicator(Document doc, XYZ point, FamilySymbol familySymbol, Workset workset)
        {
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
