using SolidWorks.Interop.sldworks;
using SolidWorks.Interop.swconst;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
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
            if (doc == null)
            {
                SwApp.SendMsgToUser2(
                    "Откройте документ перед запуском макроса.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                return;
            }

            exportPath = doc.GetPathName();
            if (string.IsNullOrWhiteSpace(exportPath))
            {
                SwApp.SendMsgToUser2(
                    "Сохраните чертёж перед запуском макроса.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                return;
            }

            int documentType = doc.GetType();
            if (documentType == (int)swDocumentTypes_e.swDocDRAWING)
            {
                ProcessSelectedTable();
                return;
            }

            if (documentType == (int)swDocumentTypes_e.swDocASSEMBLY)
            {
                ProcessAssembly();
                return;
            }

            if (documentType == (int)swDocumentTypes_e.swDocPART)
            {
                ProcessActivePart();
                return;
            }

            SwApp.SendMsgToUser2(
                "Активный документ должен быть чертежом, сборкой или деталью.",
                (int)swMessageBoxIcon_e.swMbWarning,
                (int)swMessageBoxBtn_e.swMbOk
            );
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

            string docDir = Path.GetDirectoryName(exportPath);
            string outputBaseName = GetBomOutputBaseName(bomTable);
            List<ExportRequest> exportRequests = new List<ExportRequest>();

            for (int row = 1; row < table.RowCount; row++)
            {

                int componentCount = bomTable.GetComponentsCount2(row, configuration, out string iPosition, out string iPartName);
                if (componentCount <= 0) componentCount = 1;

                if (string.IsNullOrWhiteSpace(iPartName)) continue;

                object[] components = bomTable.GetComponents(row);

                if (components == null || components.Length == 0)
                {
                    Console.WriteLine($"Ошибка: Компоненты не найдены для строки {row + 1}.");

                    IDrawingDoc drawingDoc = (IDrawingDoc)activeDoc;
                    ISheet sheet = drawingDoc.GetCurrentSheet();
                    if (sheet == null)
                    {
                        SwApp.SendMsgToUser2("Не удалось получить текущий лист чертежа.", (int)swMessageBoxIcon_e.swMbWarning, (int)swMessageBoxBtn_e.swMbOk);
                        return;
                    }

                    object[] views = sheet.GetViews();
                    if (views == null || views.Length == 0)
                    {
                        Console.WriteLine("На текущем листе не найдены виды для поиска модели.");
                        continue;
                    }

                    // Нормализуем iPartName для сравнения
                    string normalizedIPartName = iPartName?.Trim().ToLower() ?? "";
                    // Перебираем все виды, ищем совпадение по имени модели
                    foreach (IView view in views.Cast<IView>())
                    {
                        if (view == null) continue;

                        string referencedModelPath = view.GetReferencedModelName();
                        if (string.IsNullOrWhiteSpace(referencedModelPath))
                        {
                            Console.WriteLine($"Вид {view.Name} не ссылается на модель.");
                            continue;
                        }

                        string modelTitle = System.IO.Path.GetFileNameWithoutExtension(referencedModelPath)?.ToLower() ?? "";

                        // Сравниваем iPartName с названием модели
                        if (normalizedIPartName.Equals(modelTitle))
                        {
                            string configName = view.ReferencedConfiguration;

                            Console.WriteLine($"Найден совпадающий вид: {view.Name}, Модель: {referencedModelPath}");
                            AddExportRequest(exportRequests, referencedModelPath, iPartName, iPosition, componentCount, configName, docDir, outputBaseName, true);
                            break;
                        }
                        else
                        {
                            Console.WriteLine($"Вид {normalizedIPartName} не соответствует модели {modelTitle}");
                        }
                    }

                    continue;
                }

                if (components.Length > 0)
                {
                    IComponent2 component = components[0] as IComponent2;
                    if (component != null)
                    {
                        string configName = component.ReferencedConfiguration;
                        AddExportRequest(exportRequests, component.GetPathName(), iPartName, iPosition, componentCount, configName, docDir, outputBaseName, true);
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

            ProcessExportRequests(exportRequests);

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

    private void ProcessAssembly()
    {
        IModelDoc2 activeDoc = null;
        ISelectionMgr selectionMgr = null;

        try
        {
            activeDoc = SwApp.IActiveDoc2;
            selectionMgr = activeDoc?.ISelectionManager;
            List<ExportRequest> exportRequests = new List<ExportRequest>();
            string docDir = Path.GetDirectoryName(exportPath);
            string outputBaseName = SanitizeFileName(Path.GetFileNameWithoutExtension(exportPath));
            int selectedObjectCount = selectionMgr?.GetSelectedObjectCount2(-1) ?? 0;
            IAssemblyDoc assemblyDoc = activeDoc as IAssemblyDoc;
            object[] assemblyComponents = assemblyDoc?.GetComponents(false) as object[] ?? new object[0];

            if (selectedObjectCount > 0)
            {
                for (int index = 1; index <= selectedObjectCount; index++)
                {
                    IComponent2 component = selectionMgr.GetSelectedObject6(index, -1) as IComponent2;
                    if (component == null)
                    {
                        component = selectionMgr.GetSelectedObjectsComponent4(index, -1) as IComponent2;
                    }

                    if (component == null)
                    {
                        continue;
                    }

                    AddComponentExportRequest(exportRequests, component, docDir, outputBaseName, assemblyComponents);
                    Marshal.ReleaseComObject(component);
                }
            }
            else
            {
                if (assemblyComponents.Length > 0)
                {
                    foreach (IComponent2 component in assemblyComponents.Cast<IComponent2>())
                    {
                        if (component == null || component.IsSuppressed())
                        {
                            continue;
                        }

                        AddComponentExportRequest(exportRequests, component, docDir, outputBaseName, assemblyComponents);
                    }
                }
            }

            if (exportRequests.Count == 0)
            {
                SwApp.SendMsgToUser2(
                    "В сборке не найдены детали для экспорта.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                return;
            }

            ProcessExportRequests(exportRequests);

            SwApp.SendMsgToUser2(
                "Обработка компонентов завершена.",
                (int)swMessageBoxIcon_e.swMbInformation,
                (int)swMessageBoxBtn_e.swMbOk
            );
        }
        catch (Exception ex)
        {
            string errorMsg = $"Произошла ошибка при обработке сборки: {ex.Message}";
            SwApp.SendMsgToUser2(
                errorMsg,
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk
            );
            Console.WriteLine($"{errorMsg}\nСтек вызовов: {ex.StackTrace}");
        }
        finally
        {
            if (selectionMgr != null) Marshal.ReleaseComObject(selectionMgr);
            if (activeDoc != null) Marshal.ReleaseComObject(activeDoc);
        }
    }

    private void AddComponentExportRequest(
        List<ExportRequest> exportRequests,
        IComponent2 component,
        string outputRootPath,
        string outputBaseName,
        object[] assemblyComponents)
    {
        string modelPath = component.GetPathName();
        string componentName = GetComponentOutputBaseName(component);
        string configName = component.ReferencedConfiguration;
        int componentCount = CountMatchingAssemblyComponents(component, assemblyComponents);

        AddExportRequest(exportRequests, modelPath, componentName, "", componentCount, configName, outputRootPath, outputBaseName, false, true);
    }

    private int CountMatchingAssemblyComponents(IComponent2 sourceComponent, object[] assemblyComponents)
    {
        string sourcePath = sourceComponent.GetPathName();
        string sourceConfig = sourceComponent.ReferencedConfiguration ?? "";

        int count = 0;
        foreach (IComponent2 component in assemblyComponents.Cast<IComponent2>())
        {
            if (component == null || component.IsSuppressed())
            {
                continue;
            }

            if (string.Equals(component.GetPathName(), sourcePath, StringComparison.OrdinalIgnoreCase)
                && string.Equals(component.ReferencedConfiguration ?? "", sourceConfig, StringComparison.OrdinalIgnoreCase))
            {
                count++;
            }
        }

        return count > 0 ? count : 1;
    }

    private void ProcessActivePart()
    {
        IModelDoc2 activeDoc = null;
        ISelectionMgr selectionMgr = null;

        try
        {
            activeDoc = SwApp.IActiveDoc2;
            selectionMgr = activeDoc?.ISelectionManager;
            string docDir = Path.GetDirectoryName(exportPath);
            string partName = SanitizeFileName(Path.GetFileNameWithoutExtension(exportPath));
            string configuration = activeDoc.GetActiveConfiguration()?.Name;
            int selectedObjectCount = selectionMgr?.GetSelectedObjectCount2(-1) ?? 0;
            HashSet<string> selectedFlatPatternNames = GetSelectedFlatPatternNames(selectionMgr);

            if (selectedObjectCount > 0 && selectedFlatPatternNames.Count == 0)
            {
                SwApp.SendMsgToUser2(
                    "В детали выделены объекты, но среди них нет разверток FlatPattern.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                return;
            }

            List<ExportRequest> exportRequests = new List<ExportRequest>();
            AddExportRequest(exportRequests, exportPath, partName, "", 1, configuration, docDir, partName, false);

            if (exportRequests.Count > 0)
            {
                ExportRequest request = exportRequests[0];
                PrepareOutputFolders(request, new HashSet<string>(StringComparer.OrdinalIgnoreCase));

                Console.WriteLine($"\nКомпонент: {request.Position} - {request.PartName} - {request.ComponentCount}шт");
                if (selectedFlatPatternNames.Count > 0)
                {
                    Console.WriteLine($"Выбрано разверток: {selectedFlatPatternNames.Count}");
                }

                TraverseCutListFolders(
                    activeDoc,
                    request.PartName,
                    request.Position,
                    request.ComponentCount,
                    request.Configuration,
                    request.HasKnownBomPosition,
                    selectedFlatPatternNames
                );
            }

            SwApp.SendMsgToUser2(
                "Обработка детали завершена.",
                (int)swMessageBoxIcon_e.swMbInformation,
                (int)swMessageBoxBtn_e.swMbOk
            );
        }
        catch (Exception ex)
        {
            string errorMsg = $"Произошла ошибка при обработке детали: {ex.Message}";
            SwApp.SendMsgToUser2(
                errorMsg,
                (int)swMessageBoxIcon_e.swMbStop,
                (int)swMessageBoxBtn_e.swMbOk
            );
            Console.WriteLine($"{errorMsg}\nСтек вызовов: {ex.StackTrace}");
        }
        finally
        {
            if (selectionMgr != null) Marshal.ReleaseComObject(selectionMgr);
            if (activeDoc != null) Marshal.ReleaseComObject(activeDoc);
        }
    }

    private HashSet<string> GetSelectedFlatPatternNames(ISelectionMgr selectionMgr)
    {
        HashSet<string> selectedFlatPatternNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        int selectedObjectCount = selectionMgr?.GetSelectedObjectCount2(-1) ?? 0;

        for (int index = 1; index <= selectedObjectCount; index++)
        {
            Feature feature = selectionMgr.GetSelectedObject6(index, -1) as Feature;
            if (feature == null || feature.GetTypeName2() != "FlatPattern")
            {
                continue;
            }

            selectedFlatPatternNames.Add(feature.Name);
        }

        return selectedFlatPatternNames;
    }

    class ExportRequest
    {
        public string ModelPath;
        public string PartName;
        public string Position;
        public int ComponentCount;
        public string Configuration;
        public string OutputRootPath;
        public string OutputBaseName;
        public bool HasKnownBomPosition;
    }

    private void AddExportRequest(
        List<ExportRequest> exportRequests,
        string modelPath,
        string partName,
        string position,
        int componentCount,
        string configuration,
        string outputRootPath,
        string outputBaseName,
        bool hasKnownBomPosition,
        bool componentCountIsExact = false)
    {
        if (string.IsNullOrWhiteSpace(modelPath))
        {
            Console.WriteLine($"Не удалось добавить задачу экспорта для {partName}: пустой путь к модели.");
            return;
        }

        if (Path.GetExtension(modelPath).Equals(".sldprt", StringComparison.OrdinalIgnoreCase) == false)
        {
            Console.WriteLine($"Модель пропущена, так как не является деталью: {modelPath}");
            return;
        }

        if (!hasKnownBomPosition)
        {
            ExportRequest existingRequest = exportRequests.FirstOrDefault(request =>
                !request.HasKnownBomPosition
                && string.Equals(request.ModelPath, modelPath, StringComparison.OrdinalIgnoreCase)
                && string.Equals(request.Configuration ?? "", configuration ?? "", StringComparison.OrdinalIgnoreCase)
                && string.Equals(request.OutputRootPath ?? "", outputRootPath ?? "", StringComparison.OrdinalIgnoreCase)
                && string.Equals(request.OutputBaseName ?? "", outputBaseName ?? "", StringComparison.OrdinalIgnoreCase));

            if (existingRequest != null)
            {
                int normalizedComponentCount = componentCount <= 0 ? 1 : componentCount;
                existingRequest.ComponentCount = componentCountIsExact
                    ? Math.Max(existingRequest.ComponentCount, normalizedComponentCount)
                    : existingRequest.ComponentCount + normalizedComponentCount;
                return;
            }
        }

        exportRequests.Add(new ExportRequest
        {
            ModelPath = modelPath,
            PartName = partName,
            Position = position,
            ComponentCount = componentCount <= 0 ? 1 : componentCount,
            Configuration = configuration,
            OutputRootPath = outputRootPath,
            OutputBaseName = outputBaseName,
            HasKnownBomPosition = hasKnownBomPosition
        });
    }

    private string GetBomOutputBaseName(IBomTableAnnotation bomTable)
    {
        try
        {
            string referencedModelName = bomTable?.BomFeature?.GetReferencedModelName();
            string name = Path.GetFileNameWithoutExtension(referencedModelName);
            if (!string.IsNullOrWhiteSpace(name))
            {
                return SanitizeFileName(name);
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Не удалось определить имя родительской модели BOM: {ex.Message}");
        }

        return SanitizeFileName(Path.GetFileNameWithoutExtension(exportPath));
    }

    private string GetComponentOutputBaseName(IComponent2 component)
    {
        string modelName = Path.GetFileNameWithoutExtension(component.GetPathName());
        if (!string.IsNullOrWhiteSpace(modelName))
        {
            return SanitizeFileName(modelName);
        }

        return SanitizeFileName(component.Name2 ?? component.Name ?? "Component");
    }

    private string SanitizeFileName(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return "Export";
        }

        string sanitized = value;
        foreach (char invalidChar in Path.GetInvalidFileNameChars())
        {
            sanitized = sanitized.Replace(invalidChar, '_');
        }

        return sanitized.Trim();
    }

    private void PrepareOutputFolders(ExportRequest request, HashSet<string> preparedOutputFolders)
    {
        string outputBaseName = SanitizeFileName(request.OutputBaseName);
        string outputRootPath = string.IsNullOrWhiteSpace(request.OutputRootPath)
            ? Path.GetDirectoryName(exportPath)
            : request.OutputRootPath;

        DXFPath = Path.Combine(outputRootPath, $"{outputBaseName} - DXF");
        IGSPath = Path.Combine(outputRootPath, $"{outputBaseName} - IGS");

        string key = $"{DXFPath}|{IGSPath}";
        if (preparedOutputFolders.Contains(key))
        {
            return;
        }

        PrepareExportFolder(DXFPath);
        PrepareExportFolder(IGSPath);
        preparedOutputFolders.Add(key);
    }

    private void PrepareExportFolder(string folderPath)
    {
        Directory.CreateDirectory(folderPath);
    }

    private void DeleteExistingExportFile(string filePath)
    {
        if (File.Exists(filePath))
        {
            File.Delete(filePath);
        }
    }

    private void ProcessExportRequests(List<ExportRequest> exportRequests)
    {
        IModelDoc2 currentDoc = null;
        string currentModelPath = null;
        HashSet<string> preparedOutputFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        try
        {
            foreach (ExportRequest request in exportRequests)
            {
                try
                {
                    Console.WriteLine($"\nКомпонент: {request.Position} - {request.PartName} - {request.ComponentCount}шт");
                    PrepareOutputFolders(request, preparedOutputFolders);

                    if (!string.Equals(currentModelPath, request.ModelPath, StringComparison.OrdinalIgnoreCase))
                    {
                        CloseExportDocument(currentDoc);
                        currentDoc = null;
                        currentModelPath = null;

                        currentDoc = OpenAndActivateExportDocument(request.ModelPath, request.Configuration);
                        if (currentDoc == null)
                        {
                            continue;
                        }

                        currentModelPath = request.ModelPath;
                    }
                    else
                    {
                        int errors = 0;
                        IModelDoc2 activeDoc = SwApp.ActivateDoc3(
                            request.ModelPath,
                            false,
                            (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
                            ref errors
                        );

                        if (activeDoc == null || errors != 0)
                        {
                            Console.WriteLine($"Ошибка: не удалось активировать документ: {request.ModelPath}, errors={errors}");
                            continue;
                        }

                        if (!ReferenceEquals(activeDoc, currentDoc))
                        {
                            Marshal.ReleaseComObject(currentDoc);
                            currentDoc = activeDoc;
                        }
                    }

                    TraverseCutListFolders(currentDoc, request.PartName, request.Position, request.ComponentCount, request.Configuration, request.HasKnownBomPosition);
                }
                catch (Exception ex)
                {
                    string errorMsg = $"Произошла ошибка при обработке детали {request.PartName}: {ex.Message}";
                    SwApp.SendMsgToUser2(
                        errorMsg,
                        (int)swMessageBoxIcon_e.swMbStop,
                        (int)swMessageBoxBtn_e.swMbOk
                    );
                    Console.WriteLine($"{errorMsg}\nСтек вызовов: {ex.StackTrace}");
                }
            }
        }
        finally
        {
            CloseExportDocument(currentDoc);
        }
    }

    private IModelDoc2 OpenAndActivateExportDocument(string modelPath, string configuration)
    {
        int errors = 0;
        int warnings = 0;
        IModelDoc2 doc = SwApp.OpenDoc6(
            modelPath,
            (int)swDocumentTypes_e.swDocPART,
            (int)swOpenDocOptions_e.swOpenDocOptions_Silent,
            configuration ?? "",
            ref errors,
            ref warnings
        );

        if (doc == null || errors != 0)
        {
            Console.WriteLine($"Ошибка: не удалось открыть документ: {modelPath}, errors={errors}, warnings={warnings}");
            return null;
        }

        errors = 0;
        IModelDoc2 activeDoc = SwApp.ActivateDoc3(
            modelPath,
            false,
            (int)swRebuildOnActivation_e.swDontRebuildActiveDoc,
            ref errors
        );

        if (activeDoc == null || errors != 0)
        {
            Console.WriteLine($"Ошибка: не удалось активировать документ: {modelPath}, errors={errors}");
            Marshal.ReleaseComObject(doc);
            return null;
        }

        if (!ReferenceEquals(activeDoc, doc))
        {
            Marshal.ReleaseComObject(doc);
            return activeDoc;
        }

        return doc;
    }

    private void CloseExportDocument(IModelDoc2 doc)
    {
        if (doc == null)
        {
            return;
        }

        string title = doc.GetTitle();
        SwApp.CloseDoc(title);
        Marshal.ReleaseComObject(doc);
    }

    class DXFExportTask
    {
        public Feature FlatPatternFeature;
        public string FileName;
        public IModelDoc2 model;
        public string bodyName;
    }
    private void TraverseCutListFolders(
        IModelDoc2 doc,
        string iPartName,
        string iPosition,
        int iComponentCount,
        string configuration,
        bool hasKnownBomPosition,
        HashSet<string> selectedFlatPatternNames = null)
    {
        Feature feat = null;
        bool filterBySelectedFlatPatterns = selectedFlatPatternNames != null && selectedFlatPatternNames.Count > 0;

        List<DXFExportTask> exportTasks = new List<DXFExportTask>();
        Dictionary<string, int> dxfFileNameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, int> igsFileNameCounts = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        try
        {
            if (doc == null)
            {
                SwApp.SendMsgToUser2(
                    $"Не удалось получить документ детали для {iPartName}.",
                    (int)swMessageBoxIcon_e.swMbWarning,
                    (int)swMessageBoxBtn_e.swMbOk
                );
                Console.WriteLine($"Ошибка: Не удалось получить документ детали для {iPartName}.");
                return;
            }

            if (!string.IsNullOrWhiteSpace(configuration))
            {
                Console.WriteLine($"Конфигурация - '{configuration}'");

                bool success = doc.ShowConfiguration2(configuration);
                string activeConfigName = doc.GetActiveConfiguration()?.Name;

                if (activeConfigName != configuration)
                {
                    Console.WriteLine($"Ошибка: не удалось переключиться на конфигурацию - {configuration} / {success} | {doc.GetTitle()}");
                    return;
                }
            }
            else
            {
                Console.WriteLine($"Конфигурация - '{doc.GetActiveConfiguration()?.Name}'");
            }

            doc.ForceRebuild3(false);
            DisablePerspectiveView(doc);

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
                    Feature nextSubFeat = subFeat.GetNextSubFeature();
                    if (subFeat.GetTypeName2() != "CutListFolder")
                    {
                        if (subFeat != null) Marshal.ReleaseComObject(subFeat);
                        subFeat = nextSubFeat;
                        continue;
                    }

                    BodyFolder bodyFolder = (BodyFolder)subFeat.GetSpecificFeature2();
                    object[] bodies = bodyFolder?.GetBodies();
                    if (bodies == null || bodies.Length < 1)
                    {
                        if (bodyFolder != null) Marshal.ReleaseComObject(bodyFolder);
                        if (subFeat != null) Marshal.ReleaseComObject(subFeat);
                        subFeat = nextSubFeat;
                        continue;
                    }

                    cutListIndex++;


                    IBody2 firstBody = (IBody2)bodies[0];
                    if (firstBody == null)
                    {
                        if (bodyFolder != null) Marshal.ReleaseComObject(bodyFolder);
                        if (subFeat != null) Marshal.ReleaseComObject(subFeat);
                        subFeat = nextSubFeat;
                        continue;
                    }




                    int cutListType = bodyFolder.GetCutListType();
                    /*
                        -1 — Неизвестный тип (swCutListType_Unknown)

                        1 — Твердотельное тело (swCutListType_SolidBody)

                        2 — Листовой металл (swCutListType_Sheetmetal)

                        3 — Сварная конструкция (swCutListType_Weldment)
                    */


                    if (cutListType == 3) // Сварная конструкция
                    {
                        if (filterBySelectedFlatPatterns)
                        {
                            Marshal.ReleaseComObject(firstBody);
                            if (bodyFolder != null) Marshal.ReleaseComObject(bodyFolder);
                            if (subFeat != null) Marshal.ReleaseComObject(subFeat);
                            subFeat = nextSubFeat;
                            continue;
                        }

                        var (length, width, depth) = GetWeldBodyProperties(subFeat);

                        string fileName = BuildExportFileName(iPosition, cutListIndex, iPartName, configuration, $"{width}х{depth}х{length}", bodies.Length * iComponentCount, hasKnownBomPosition);
                        fileName = BuildUniqueExportFileName(fileName, igsFileNameCounts);

                        ExportBodyToIGS(firstBody, fileName);
                    }

                    if (cutListType == 2) // Листовой металл
                    {
                        var (length, width, thickness) = GetSheetMetalProperties(subFeat); ;

                        string fileName = BuildExportFileName(iPosition, cutListIndex, iPartName, configuration, $"{length}х{width}х{thickness}", bodies.Length * iComponentCount, hasKnownBomPosition);

                        object[] bodyFeatures = firstBody.GetFeatures() as object[];
                        if (bodyFeatures == null || bodyFeatures.Length < 1)
                        {
                            Console.WriteLine($"Не найдены фичи тела для экспорта DXF: {firstBody.Name}");
                        }
                        else
                        {

                            foreach (IFeature bodyFeat in bodyFeatures.Cast<IFeature>())
                            {
                                string bodyFeatType = bodyFeat.GetTypeName2();
                                if (bodyFeatType != "FlatPattern") continue;
                                if (filterBySelectedFlatPatterns && !selectedFlatPatternNames.Contains(bodyFeat.Name)) continue;

                                string uniqueFileName = BuildUniqueExportFileName(fileName, dxfFileNameCounts);

                                // Добавляем задание на экспорт
                                exportTasks.Add(new DXFExportTask
                                {
                                    FlatPatternFeature = (Feature)bodyFeat,
                                    FileName = uniqueFileName,
                                    model = doc,
                                    bodyName = firstBody.Name
                                });
                            }
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
            if (feat != null) Marshal.ReleaseComObject(feat);
        }
    }

    private string BuildExportFileName(
        string position,
        int cutListIndex,
        string partName,
        string configuration,
        string dimensions,
        int quantity,
        bool hasKnownBomPosition)
    {
        string configurationSuffix = string.IsNullOrWhiteSpace(configuration)
            ? ""
            : $" - {SanitizeFileName(configuration)}";

        string positionPrefix = hasKnownBomPosition
            ? $"{position}.{cutListIndex} - "
            : "";

        return $"{positionPrefix}{partName.Trim()}{configurationSuffix}  {dimensions}мм - {quantity}шт";
    }

    private string BuildUniqueExportFileName(string fileName, Dictionary<string, int> fileNameCounts)
    {
        if (!fileNameCounts.TryGetValue(fileName, out int count))
        {
            fileNameCounts[fileName] = 1;
            return fileName;
        }

        count++;
        fileNameCounts[fileName] = count;
        return $"{fileName} - {count}";
    }

    private void DisablePerspectiveView(IModelDoc2 doc)
    {
        IModelView activeView = doc?.IActiveView;
        if (activeView != null && activeView.HasPerspective())
        {
            activeView.RemovePerspective();
            Console.WriteLine("Перспектива отключена для экспорта.");
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

        double[] dataAlignment = BuildFlatPatternAlignment(flatPattern, swModel);

        object varAlignment;

        varAlignment = dataAlignment;

        string filePath = Path.Combine(DXFPath, $"{fileName}.dxf");
        DeleteExistingExportFile(filePath);
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

    private double[] BuildFlatPatternAlignment(Feature flatPattern, IModelDoc2 swModel)
    {
        double[] normal = GetFlatPatternFixedFaceNormal(flatPattern, swModel);
        if (normal == null || !NormalizeVector(normal))
        {
            normal = new double[] { 0, 0, 1 };
        }

        double[] xDirection = ProjectVectorToPlane(new double[] { 1, 0, 0 }, normal);
        if (!NormalizeVector(xDirection))
        {
            xDirection = ProjectVectorToPlane(new double[] { 0, 1, 0 }, normal);
            NormalizeVector(xDirection);
        }

        double[] yDirection = CrossProduct(normal, xDirection);
        NormalizeVector(yDirection);

        return new double[]
        {
            xDirection[0], xDirection[1], xDirection[2],
            yDirection[0], yDirection[1], yDirection[2],
            normal[0], normal[1], normal[2],
            0, 0, 0
        };
    }

    private double[] GetFlatPatternFixedFaceNormal(Feature flatPattern, IModelDoc2 swModel)
    {
        IFlatPatternFeatureData flatPatternData = null;

        try
        {
            flatPatternData = flatPattern.GetDefinition() as IFlatPatternFeatureData;
            if (flatPatternData == null)
            {
                return null;
            }

            flatPatternData.AccessSelections(swModel, null);
            IFace2 fixedFace = flatPatternData.FixedFace2 as IFace2;
            if (fixedFace == null)
            {
                return null;
            }

            object normalObject = fixedFace.Normal;
            if (normalObject is double[] normal && normal.Length >= 3)
            {
                return new double[] { normal[0], normal[1], normal[2] };
            }

            ISurface surface = fixedFace.GetSurface() as ISurface;
            if (surface != null && surface.IsPlane())
            {
                object planeParamsObject = surface.PlaneParams;
                if (planeParamsObject is double[] planeParams && planeParams.Length >= 6)
                {
                    return new double[] { planeParams[3], planeParams[4], planeParams[5] };
                }
            }

            return null;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Не удалось получить нормаль фиксированной грани развертки: {ex.Message}");
            return null;
        }
        finally
        {
            if (flatPatternData != null)
            {
                flatPatternData.ReleaseSelectionAccess();
            }
        }
    }

    private double[] ProjectVectorToPlane(double[] vector, double[] normal)
    {
        double dot = DotProduct(vector, normal);
        return new double[]
        {
            vector[0] - dot * normal[0],
            vector[1] - dot * normal[1],
            vector[2] - dot * normal[2]
        };
    }

    private bool NormalizeVector(double[] vector)
    {
        double length = Math.Sqrt(DotProduct(vector, vector));
        if (length < 1e-9)
        {
            return false;
        }

        vector[0] /= length;
        vector[1] /= length;
        vector[2] /= length;
        return true;
    }

    private double DotProduct(double[] first, double[] second)
    {
        return first[0] * second[0] + first[1] * second[1] + first[2] * second[2];
    }

    private double[] CrossProduct(double[] first, double[] second)
    {
        return new double[]
        {
            first[1] * second[2] - first[2] * second[1],
            first[2] * second[0] - first[0] * second[2],
            first[0] * second[1] - first[1] * second[0]
        };
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
            DeleteExistingExportFile(fullPath);

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
