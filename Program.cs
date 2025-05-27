
using Microsoft.SqlServer.Server;
using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.RegularExpressions;

public class SolidWorksMacro
{

    public static ISldWorks SwApp;
    String exportPath = "";
    String DXFPath = "";
    String IGSPath = "";
    public static void Main()
    {
        SolidWorksMacro macro = new SolidWorksMacro();
        macro.Run();
    }

    private void Run()
    {
        try
        {
            SwApp = (ISldWorks)Marshal.GetActiveObject("SldWorks.Application");
            if (SwApp == null)
            {
                throw new Exception("Не удалось подключиться к SolidWorks. Убедитесь, что SolidWorks запущен.");
            }

            IModelDoc2 doc = SwApp.IActiveDoc2;
            exportPath = doc.GetPathName();
            if (doc == null)
            {
                SwApp.SendMsgToUser2(
                    "Откройте документ перед запуском макроса.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                return;
            }

            if (doc.GetType() != (int)swDocumentTypes_e.swDocDRAWING)
            {
                SwApp.SendMsgToUser2(
                    "Активный документ должен быть чертежом!",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                return;
            }

            ProcessSelectedTable();
        }
        catch (Exception ex)
        {
            string errorMsg = $"Произошла ошибка: {ex.Message}";
            if (SwApp != null)
            {
                SwApp.SendMsgToUser2(
                    errorMsg,
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk
                );
            }
            Console.WriteLine($"{errorMsg}\nСтек вызовов: {ex.StackTrace}");
        }
        finally
        {
            if (SwApp != null)
            {
                Marshal.ReleaseComObject(SwApp);
                SwApp = null;
            }
        }
    }

    private void ProcessSelectedTable()
    {
        IModelDoc2 activeDoc = null;
        ISelectionMgr selectionMgr = null;
        ITableAnnotation table = null;
        IBomTableAnnotation bomTable = null;
        string configuration = "";

        try
        {
            activeDoc = SwApp.IActiveDoc2;
            if (activeDoc == null)
            {
                SwApp.SendMsgToUser2(
                    "Активный документ не найден.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Активный документ не найден.");
                return;
            }

            selectionMgr = activeDoc.ISelectionManager;
            if (selectionMgr == null)
            {
                SwApp.SendMsgToUser2(
                    "Не удалось получить SelectionManager.",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Не удалось получить SelectionManager.");
                return;
            }

            object selectedObject = selectionMgr.GetSelectedObject6(1, -1);
            if (selectedObject == null)
            {
                SwApp.SendMsgToUser2(
                    "Выделенный объект не найден. Пожалуйста, выберите таблицу в дереве проекта.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Выделенный объект не найден.");
                return;
            }


            table = selectedObject as ITableAnnotation;
            if (table == null)
            {
                SwApp.SendMsgToUser2(
                    "Выделенный объект не является таблицей. Пожалуйста, выберите таблицу.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Выделенный объект не является таблицей.");
                return;
            }

            bomTable = table as IBomTableAnnotation;

            if (bomTable == null)
            {
                SwApp.SendMsgToUser2(
                    "Выделенная таблица не является BOM-таблицей.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Выделенная таблица не является BOM-таблицей.");
                return;
            }

            Console.WriteLine($"Обработка таблицы: {table.RowCount} строк");

            // Создать папку, если нужно
            string docDir = Path.GetDirectoryName(exportPath);
            DXFPath = Path.Combine(docDir, "DXF");
            IGSPath = Path.Combine(docDir, "IGS");

            if (Directory.Exists(DXFPath))
            {
                // Удаляем все файлы в папке
                foreach (string file in Directory.GetFiles(DXFPath))
                {
                    File.Delete(file);
                }
            }
            else
            {
                // Создаем папку, если она не существует
                Directory.CreateDirectory(DXFPath);
            }

            if (Directory.Exists(IGSPath))
            {
                // Удаляем все файлы в папке
                foreach (string file in Directory.GetFiles(IGSPath))
                {
                    File.Delete(file);
                }
            }
            else
            {
                // Создаем папку, если она не существует
                Directory.CreateDirectory(IGSPath);
            }

            for (int row = 1; row < table.RowCount; row++)
            {

                int componentCount = bomTable.GetComponentsCount2(row, configuration, out string iPosition, out string iPartName);

                if (string.IsNullOrWhiteSpace(iPartName)) continue;

                Console.WriteLine($"\nКомпонент: {iPosition} - {iPartName} - {componentCount}шт");

                object[] components = bomTable.GetComponents(row);
                
                if (components == null || components.Length < 0)
                {
                    Console.WriteLine($"Ошибка: Компоненты не найдены для строки {row + 1}.");

                    IDrawingDoc drawingDoc = (IDrawingDoc)activeDoc;
                    ISheet sheet = drawingDoc.GetCurrentSheet();
                    if (sheet == null)
                    {
                        SwApp.SendMsgToUser2("Не удалось получить текущий лист чертежа.", (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk);
                        return;
                    }

                    IModelDoc2 model = null;

                    object[] views = sheet.GetViews();
                    // Нормализуем iPartName для сравнения
                    string normalizedIPartName = iPartName?.Trim().ToLower() ?? "";
                    // Перебираем все виды, ищем совпадение по имени модели
                    foreach (IView view in views.Cast<IView>())
                    {
                        if (view == null) continue;

                        model = view.ReferencedDocument;
                        if (model == null)
                        {
                            Console.WriteLine($"Вид {view.Name} не ссылается на модель.");
                            continue;
                        }

                        string modelTitle = System.IO.Path.GetFileNameWithoutExtension(model.GetTitle())?.ToLower() ?? "";

                        // Сравниваем iPartName с названием модели
                        if (normalizedIPartName.Equals(modelTitle))
                        {
                            string configName = view.ReferencedConfiguration;

                            Console.WriteLine($"Найден совпадающий вид: {view.Name}, Модель: {model.GetTitle()}");
                            if (model == null)
                            {
                                SwApp.SendMsgToUser2("Не удалось получить ссылочную модель из вида.", (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk);
                                return;
                            }

                            // Проверка, является ли модель деталью
                            if (model.GetType() == (int)swDocumentTypes_e.swDocPART)
                            {
                                IPartDoc partDoc = (IPartDoc)model;
                                Console.WriteLine($"Деталь, связанная с таблицей: {model.GetTitle()}");

                                // Используем количество из таблицы, если доступно
                                int useComponentCount = 1;

                                if (bomTable != null)
                                {
                                    useComponentCount = bomTable.GetComponentsCount2(row, configuration, out _, out _);
                                    if (useComponentCount <= 0) useComponentCount = 1;
                                }

                                TraverseCutListFolders(partDoc, iPartName, iPosition, useComponentCount, configName);
                            }
                            break;
                        }
                        else
                        {
                            Console.WriteLine($"Вид {normalizedIPartName} не соответствует модели {modelTitle}");
                        }

                        // Освобождаем временный model, если он не подходит
                        Marshal.ReleaseComObject(model);
                        model = null;
                    }

                    continue;
                }

                if (components.Length > 0)
                {
                    IComponent2 component = components[0] as IComponent2;
                    if (component != null)
                    {
                        IModelDoc2 model = component.GetModelDoc2();
                        if (model == null)
                        {
                            Console.WriteLine($"Не удалось получить модель для компонента в строке {row + 1}.");
                            continue;
                        }

                        string configName = component.ReferencedConfiguration;
                        // Console.WriteLine($"Конфигурация компонента: {configName}");

                        int modelType = model.GetType();
                        if (modelType == (int)swDocumentTypes_e.swDocPART)
                        {
                            IPartDoc partDoc = (IPartDoc)model;
                            TraverseCutListFolders(partDoc, iPartName, iPosition, componentCount, configName);
                        }
                        Marshal.ReleaseComObject(model);
                        Marshal.ReleaseComObject(component);
                    }
                    else
                    {
                        SwApp.SendMsgToUser2(
                            $"Не удалось получить компонент для строки {row + 1}.",
                            (int)swMessageBoxIcon_e.swMbWarning,
                            (int)swMessageBoxBtn_e.swMbOk
                        );
                        Console.WriteLine($"Ошибка: Не удалось получить компонент для строки {row + 1}.");
                    }
                }
            }

            SwApp.SendMsgToUser2(
                "Обработка таблицы завершена.",
                (int)swMessageBoxIcon_e.swMbInformation,
                (int)swMessageBoxBtn_e.swMbOk
            );
            Console.WriteLine("\nОбработка таблицы завершена.");
        }
        catch (Exception ex)
        {
            string errorMsg = $"Произошла ошибка при обработке таблицы: {ex.Message}";
            SwApp.SendMsgToUser2(
                errorMsg,
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk
            );
            Console.WriteLine($"{errorMsg}\nСтек вызовов: {ex.StackTrace}");
        }
        finally
        {
            if (bomTable != null) Marshal.ReleaseComObject(bomTable);
            if (table != null) Marshal.ReleaseComObject(table);
            if (selectionMgr != null) Marshal.ReleaseComObject(selectionMgr);
            if (activeDoc != null) Marshal.ReleaseComObject(activeDoc);
        }
    }

    class DXFExportTask
    {
        public Feature FlatPatternFeature;
        public string FileName;
        public IModelDoc2 model;
        public string bodyName;
    }
    private void TraverseCutListFolders(IPartDoc partDoc, string iPartName, string iPosition, int iComponentCount, string configuration)
    {
        IModelDoc2 model = null;
        Feature feat = null;

        List<DXFExportTask> exportTasks = new List<DXFExportTask>();
        try
        {
            if (partDoc == null)
            {
                SwApp.SendMsgToUser2(
                    $"Не удалось получить документ детали для {iPartName}.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine($"Ошибка: Не удалось получить документ детали для {iPartName}.");
                return;
            }
            
            model = (IModelDoc2)partDoc;
            string modelPath = model.GetPathName();

            // Активируем документ
            int errors = 0;
            IModelDoc2 doc = SwApp.ActivateDoc3(
                modelPath,
                false,
                (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
                ref errors
            );

            if (errors != 0)
            {
                SwApp.SendMsgToUser2(
                    $"Не удалось активировать документ: {modelPath}",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine($"Ошибка: Не удалось активировать документ: {modelPath}");
                return;
            }

            
            bool success = doc.ShowConfiguration2(configuration);
            string activeConfigName = doc.GetActiveConfiguration()?.Name;
            if (activeConfigName != configuration)
            {
                Console.WriteLine($"Ошибка: не удалось переключиться на конфигурацию - {configuration} / {success} | {doc.GetTitle()}");
                return;
            }

            feat = doc.FirstFeature();
            Feature nextFeat = null;
            while (feat != null)
            {
                string featType = feat.GetTypeName2();
                if (featType != "SolidBodyFolder")
                {
                    nextFeat = feat.GetNextFeature();
                    Marshal.ReleaseComObject(feat);
                    feat = nextFeat;
                    continue;
                }

                // Console.WriteLine($"Найдена папка Solid Bodies: {iPosition}");

                Feature subFeat = feat.GetFirstSubFeature();
                int cutListIndex = 0;
                while (subFeat != null)
                {
                    if (subFeat.GetTypeName2() != "CutListFolder") continue;
                        
                    BodyFolder bodyFolder = (BodyFolder)subFeat.GetSpecificFeature2();
                    object[] bodies = bodyFolder?.GetBodies();
                    if (bodies == null || bodies.Length < 1)
                    {
                        Marshal.ReleaseComObject(bodyFolder);
                        foreach (object objBody in bodies)
                        {
                            Marshal.ReleaseComObject(objBody);
                        }
                        continue;
                    }

                    cutListIndex++;
                    

                    IBody2 firstBody = (IBody2)bodies[0];
                    if (firstBody == null) continue;

                    
                    

                    int cutListType = bodyFolder.GetCutListType();
                    /*
                        -1 — Неизвестный тип (swCutListType_Unknown)

                        1 — Твердотельное тело (swCutListType_SolidBody)

                        2 — Листовой металл (swCutListType_Sheetmetal)

                        3 — Сварная конструкция (swCutListType_Weldment)
                    */
                    

                    if (cutListType == 3) // Сварная конструкция
                    {
                        var (length, width, depth ) = GetWeldBodyProperties(subFeat);

                        string fileName = $"{iPosition}.{cutListIndex} - {iPartName.Trim()}  {width}х{depth}х{length}мм - {bodies.Length * iComponentCount}шт";

                        ExportBodyToIGS(firstBody, fileName);
                    }

                    if (cutListType == 2) // Листовой металл
                    {
                        var (length, width, thickness) = GetSheetMetalProperties(subFeat);;

                        string fileName = $"{iPosition}.{cutListIndex} - {iPartName.Trim()}  {length}х{width}х{thickness}мм - {bodies.Length * iComponentCount}шт";

                        object[] bodyFeatures = firstBody.GetFeatures() as object[];
                        if (bodyFeatures.Length < 1) return;

                        foreach (IFeature bodyFeat in bodyFeatures.Cast<IFeature>())
                        {
                            string bodyFeatType = bodyFeat.GetTypeName2();
                            if (bodyFeatType != "FlatPattern") continue;

                            // Добавляем задание на экспорт
                            exportTasks.Add(new DXFExportTask
                            {
                                FlatPatternFeature = (Feature)bodyFeat,
                                FileName = fileName,
                                model = doc,
                                bodyName = firstBody.Name
                            });
                        }
                    }


                    Marshal.ReleaseComObject(firstBody);
                    if (bodies == null || bodies.Length < 1)
                    {
                        Marshal.ReleaseComObject(bodyFolder);
                        foreach (object objBody in bodies)
                        {
                            if (objBody is IBody2 body)
                            {
                                Marshal.ReleaseComObject(body);
                            }
                        }
                    }

                    if (bodyFolder != null) Marshal.ReleaseComObject(bodyFolder);
                        

                    Feature nextSubFeat = subFeat.GetNextSubFeature();
                    if (subFeat != null) Marshal.ReleaseComObject(subFeat);
                    subFeat = nextSubFeat;
                }

                nextFeat = feat.GetNextFeature();
                if (feat != null) Marshal.ReleaseComObject(feat);
                feat = nextFeat;
            }

            // Console.WriteLine($"    Заданий всего: {exportTasks.LongCount()}");
            foreach (var task in exportTasks)
            {
                try
                {
                    ExportSheetMetalToDXF(task.FlatPatternFeature, task.FileName, task.model);
                    
                    // ExportFlatPatternToDXF(task.FlatPatternFeature, task.bodyName, task.FileName, task.model); 
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Ошибка при экспорте {task.FileName}: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            string errorMsg = $"Произошла ошибка при обработке Cut List для {iPartName}: {ex.Message}";
            SwApp.SendMsgToUser2(
                errorMsg,
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk
            );
            Console.WriteLine($"{errorMsg}\nСтек вызовов: {ex.StackTrace}");
        }
        finally
        {
            // Проверяешь, остался ли активным нужный документ
            IModelDoc2 currentActiveDoc = SwApp.IActiveDoc2;
            if (currentActiveDoc != null)
            {
                string currentPath = currentActiveDoc.GetPathName();
                if (currentPath == model.GetPathName())
                {
                    // Документ task.model все еще активен
                    SwApp.CloseDoc(model.GetTitle());
                }
            }
            Marshal.ReleaseComObject(currentActiveDoc);
            if (model != null) Marshal.ReleaseComObject(model);
            if (feat != null) Marshal.ReleaseComObject(feat);
        }
    }

    public bool ExportFlatPatternToDXF(Feature FlatPatternFeature, string bodyName, string fileName, IModelDoc2 swModel)
    {

        ModelDoc2 drawingDoc = null;
        int errors = 0;

        try
        {
            // Получаем тело развертки
            

            // Скрываем все тела кроме нужного (временно)

            // Создаем новый чертеж
            drawingDoc = (ModelDoc2)SwApp.NewDocument(
                SwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplateDrawing),
                (int)swDwgPaperSizes_e.swDwgPaperA4size, 0, 0);

            if (drawingDoc == null)
                throw new Exception("Не удалось создать новый чертеж");

            DrawingDoc swDraw = (DrawingDoc)drawingDoc;
            Sheet currentSheet = swDraw.GetCurrentSheet();
            currentSheet.SetTemplateName(""); // сбросить шаблон
            currentSheet.SetSheetFormatName(""); // убрать формат

            // Добавляем вид модели на чертеж
            string modelPath = swModel.GetPathName();
            string flatName = FlatPatternFeature.Name;

            // Добавляем вид спереди (или развертку)
            View baseView = swDraw.CreateDrawViewFromModelView3(
                modelPath, "", 0, 0, 0);
            if (baseView == null)
                throw new Exception("Не удалось создать вид развертки на чертеже");

            baseView.ScaleDecimal = 1.0; // Устанавливаем масштаб 1:1
            
            object[] arrBody = (object[])baseView.Bodies;
            List<Body2> arrBodiesIn = new List<Body2>();

            for (int i = 0; i < arrBody.Length; i++)
            {
                Body2 swBody = (Body2)arrBody[i];
                Console.WriteLine($"Проверяем тело: {swBody.Name} {bodyName} {flatName}");
                if (swBody.Name == bodyName)
                {
                    arrBodiesIn.Add(swBody); // Добавляем нужное тело в список
                }
                else
                {
                    Marshal.ReleaseComObject(swBody); // Освобождаем не нужные тела
                }
            }
            // Если дальше нужен именно object[]
            object[] bodiesArray = arrBodiesIn.ToArray();

            // Сохраняем чертеж в DXF
            string filePath = Path.Combine(DXFPath, $"{fileName}.dxf");
            bool saved = drawingDoc.Extension.SaveAs(
                filePath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null, ref errors, ref errors);

            if (!saved)
                throw new Exception("Не удалось сохранить DXF. Код ошибки: " + errors);

            return true;
        }
        catch (Exception ex)
        {
            SwApp.SendMsgToUser2("Ошибка экспорта развертки: " + ex.Message,
                (int)swMessageBoxIcon_e.swMbStop, (int)swMessageBoxBtn_e.swMbOk);
            return false;
        }
        finally
        {
            
            // Закрываем чертеж без сохранения, если не нужно хранить его
            if (drawingDoc != null)
            {
                string docName = drawingDoc.GetTitle();
                SwApp.CloseDoc(docName);
            }
        }
    }

    public void ExportSheetMetalToDXF(
        Feature flatPattern,
        string fileName,
        IModelDoc2 swModel
    )
    {

        int sheetMetalOptions = BuildSheetMetalOptions(
            exportGeometry: true,           // Экспортировать геометрию плоского шаблона (бит 1)
            includeHiddenEdges: false,      // Включать скрытые кромки (бит 2)
            exportBendLines: true,          // Экспортировать линии сгиба (бит 3)
            includeSketches: false,         // Включать эскизы (бит 4)
            mergeCoplanarFaces: false,      // Объединять копланарные грани (бит 5)
            exportLibraryFeatures: false,   // Экспортировать библиотечные элементы (бит 6)
            exportFormingTools: false,      // Экспортировать формообразующие инструменты (бит 7)
            exportBoundingBox: false        // Экспортировать габаритный прямоугольник (бит 12)
        );

        PartDoc partDoc = swModel as PartDoc;

        double[] dataAlignment = new double[12];
        object varAlignment;

        dataAlignment[0] = 0.0;
        dataAlignment[1] = 0.0;
        dataAlignment[2] = 0.0;
        dataAlignment[3] = 0.0;
        dataAlignment[4] = 0.0;
        dataAlignment[5] = 0.0;
        dataAlignment[6] = 0.0;
        dataAlignment[7] = 0.0;
        dataAlignment[8] = 0.0;
        dataAlignment[9] = 0.0;
        dataAlignment[10] = 0.0;
        dataAlignment[11] = 0.0;

        varAlignment = dataAlignment;

        string filePath = Path.Combine(DXFPath, $"{fileName}.dxf");
        // Console.WriteLine($"    DXF: {filePath} ");
        // Выбор фичи и экспорт
        if (flatPattern.Select2(false, -1))
        {
            bool result = partDoc.ExportToDWG2(
                filePath,                                           // Путь к файлу для сохранения экспортированного DWG
                swModel.GetPathName(),                              // Путь к исходной модели SolidWorks
                (int)swExportToDWG_e.swExportToDWG_ExportSheetMetal,// Режим экспорта: экспорт листового металла
                true,                                               // Экспортировать плоский шаблон (развертку)
                varAlignment,                                       // Массив из 12 значений double, содержащий информацию, связанную с выравниванием выходных данных
                false,                                              // Экспортировать только выбранные элементы (false = все элементы)
                false,                                              // Игнорировать невидимые слои (false = включать все слои)
                sheetMetalOptions,                                  // Битовая маска опций для экспорта листового металла
                null                                                // Массив имен представлений аннотаций для экспорта
            );

            if (!result)
            {
                throw new Exception("Failed to export flat pattern");
            }
            else
            {
                Console.WriteLine($"    DXF Успешно сохранен: {filePath}");
            }
        }
    }

    /// <summary>
    /// Собирает битовую маску SheetMetalOptions для ExportToDWG2.
    /// </summary>
    public int BuildSheetMetalOptions(
        bool exportGeometry = true,
        bool includeHiddenEdges = false,
        bool exportBendLines = true,
        bool includeSketches = false,
        bool mergeCoplanarFaces = false,
        bool exportLibraryFeatures = false,
        bool exportFormingTools = false,
        bool exportBoundingBox = false
    )
    {
        int options = 0;

        if (exportGeometry) options |= 1 << 0;  // Bit 1
        if (includeHiddenEdges) options |= 1 << 1;  // Bit 2
        if (exportBendLines) options |= 1 << 2;  // Bit 3
        if (includeSketches) options |= 1 << 3;  // Bit 4
        if (mergeCoplanarFaces) options |= 1 << 4;  // Bit 5
        if (exportLibraryFeatures) options |= 1 << 5;  // Bit 6
        if (exportFormingTools) options |= 1 << 6;  // Bit 7
        if (exportBoundingBox) options |= 1 << 11; // Bit 12 (обрати внимание — это 2^11)

        return options;
    }

    private void ExportBodyToIGS(IBody2 body, string fileName)
    {
        IModelDoc2 activeDoc = null;
        IModelDoc2 newPart = null;
        IPartDoc partDoc = null;
        Feature feature = null;

        try
        {
            if (body == null || string.IsNullOrWhiteSpace(fileName))
            {
                SwApp.SendMsgToUser2(
                    "Некорректные параметры для экспорта тела.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Некорректные параметры для экспорта тела.");
                return;
            }



            int errors = 0;
            int warnings = 0;
            newPart = SwApp.NewDocument(
                SwApp.GetUserPreferenceStringValue((int)swUserPreferenceStringValue_e.swDefaultTemplatePart),
                (int)swDwgPaperSizes_e.swDwgPaperA4size,
                0, 0
            );

            if (newPart == null)
            {
                SwApp.SendMsgToUser2(
                    "Не удалось создать новый документ детали.",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Не удалось создать новый документ детали.");
                return;
            }

            partDoc = newPart as IPartDoc;
            if (partDoc == null)
            {
                SwApp.SendMsgToUser2(
                    "Не удалось привести документ к типу PartDoc.",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Не удалось привести документ к типу PartDoc.");
                return;
            }

            feature = partDoc.CreateFeatureFromBody3(body, false, 0);
            if (feature == null)
            {
                SwApp.SendMsgToUser2(
                    "Не удалось создать фичу из тела.",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Не удалось создать фичу из тела.");
                SwApp.CloseDoc(newPart.GetTitle());
                return;
            }

            if (string.IsNullOrWhiteSpace(exportPath))
            {
                SwApp.SendMsgToUser2(
                    "Документ не сохранён. Сохраните документ перед экспортом.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine("Ошибка: Документ не сохранён.");
                return;
            }

            string fullPath = Path.Combine(IGSPath, $"{fileName}.igs");

            bool saveResult = newPart.Extension.SaveAs(
                fullPath,
                (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
                (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
                null,
                ref errors,
                ref warnings
            );

            if (!saveResult || errors != 0)
            {
                SwApp.SendMsgToUser2(
                    $"Ошибка при сохранении в IGES: код ошибки {errors}, предупреждения: {warnings}",
                    (int)swMessageBoxIcon_e.swMbStop,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine($"Ошибка при сохранении в IGES: код ошибки {errors}, предупреждения: {warnings}");
            }
            else
            {
                Console.WriteLine($"    IGES Успешно сохранен: {fullPath}");
            }
        }
        catch (Exception ex)
        {
            string errorMsg = $"Произошла ошибка при экспорте тела в IGES: {ex.Message}";
            SwApp.SendMsgToUser2(
                errorMsg,
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk
            );
            Console.WriteLine($"{errorMsg}\nСтек вызовов: {ex.StackTrace}");
        }
        finally
        {
            if (newPart != null)
            {
                SwApp.CloseDoc(newPart.GetTitle());
                Marshal.ReleaseComObject(newPart);
            }
            if (partDoc != null) Marshal.ReleaseComObject(partDoc);
            if (feature != null) Marshal.ReleaseComObject(feature);
            if (activeDoc != null) Marshal.ReleaseComObject(activeDoc);
        }
    }

    private (string Length, string Width, string Depth) GetWeldBodyProperties(Feature feat)
    {
        CustomPropertyManager propMgr = feat.CustomPropertyManager;

        string[] lengthKeys = { "Длина", "Length" };
        string[] descriptionKeys = { "Описание", "Description" };

        string length = FindPropertyValue(propMgr, lengthKeys);
        string description = FindPropertyValue(propMgr, descriptionKeys);

        string cleaned = description.Replace(" ", "");

        string width = "-";
        string depth = "-";

        if (!string.IsNullOrWhiteSpace(description))
        {
            // Ищем шаблон вроде "40,00 X 40,00 X 2,00"
            Regex regex = new Regex(@"(\d{1,4}(?:[.,]\d{1,2})?)[xX×](\d{1,4}(?:[.,]\d{1,2})?)[xX×](\d{1,4}(?:[.,]\d{1,2})?)");

            Match match = regex.Match(cleaned);

            if (match.Success)
            {
                width = match.Groups[1].Value;
                depth = match.Groups[2].Value;
                string thickness = match.Groups[3].Value;

                // Console.WriteLine($"Ширина: {width}, Высота: {depth}, Толщина: {thickness}");
            }
            else
            {
                Console.WriteLine("⚠ Не удалось распарсить габариты профиля из описания.");
            }
        }

        return (length, width, depth);
    }


    private (string Length, string Width, string Thickness) GetSheetMetalProperties(Feature feet)
    {
        CustomPropertyManager propMgr = feet.CustomPropertyManager;

        string[] lengthKeys = new[] { "Длина граничной рамки", "Bounding Box Length" };
        string[] widthKeys = new[] { "Ширина граничной рамки", "Bounding Box Width" };
        string[] thicknessKeys = new[] { "Толщина листового металла", "Sheet Metal Thickness" };

        string lengthRaw = FindPropertyValue(propMgr, lengthKeys);
        string widthRaw = FindPropertyValue(propMgr, widthKeys);
        string thickness = FindPropertyValue(propMgr, thicknessKeys);

        string FormatToSingleDecimal(string input)
        {
            if (double.TryParse(input.Replace(',', '.'), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out double value))
            {
                return value.ToString("F1", System.Globalization.CultureInfo.InvariantCulture);
            }
            return input; // Возвращаем как есть, если не получилось распарсить
        }

        string length = FormatToSingleDecimal(lengthRaw);
        string width = FormatToSingleDecimal(widthRaw);

        return (length, width, thickness);
    }


    private string FindPropertyValue(ICustomPropertyManager propMgr, string[] possibleKeys)
    {
        foreach (string key in possibleKeys)
        {
            bool found = propMgr.Get4(key, false, out string valOut, out string resolvedVal);
            if (found && !string.IsNullOrWhiteSpace(resolvedVal))
            {
                // Console.WriteLine($"Найдено свойство: {key} = {resolvedVal}");
                return resolvedVal;
            }
        }

        Console.WriteLine("Свойство не найдено ни по одному из ключей: " + string.Join(", ", possibleKeys));
        return "-";
    }

}