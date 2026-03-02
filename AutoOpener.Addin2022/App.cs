using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using AutoOpener.Core.IO;
using AutoOpener.Core.Jobs;
using AutoOpener.Core.Models;
using AutoOpener.Core.Processes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;

namespace AutoOpener.Addin2022
{
    public class App : IExternalApplication
    {
        private const int ThisVersion = 2022;

        private ExternalEvent _extEvent;
        private OpenJobHandler _handler;
        private FileSystemWatcher _watcher;
        private System.Timers.Timer _recoveryTimer;
        private static readonly int _pid = Process.GetCurrentProcess().Id;

        public Result OnStartup(UIControlledApplication application)
        {
            PathsService.SetVersionContext(ThisVersion);
            Directory.CreateDirectory(PathsService.LogsDir);
            Directory.CreateDirectory(PathsService.OutDir);
            CleanupService.CleanupOldArtifacts(ThisVersion, 7);
            Logger.Info("Addin 2022 startup");

            _handler = new OpenJobHandler(ThisVersion);
            _extEvent = ExternalEvent.Create(_handler);

            application.DialogBoxShowing += OnDialogBoxShowing;

            try
            {
                _watcher = new FileSystemWatcher(PathsService.QueueDirFor(ThisVersion));
                _watcher.Filter = "*.*";
                _watcher.Created += OnFileCreated;
                _watcher.EnableRaisingEvents = true;

                _extEvent.Raise();
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to start FileSystemWatcher: " + ex);
                
                // Инициализация таймера восстановления зависших задач
                _recoveryTimer = new System.Timers.Timer(300000); // 5 минут
                _recoveryTimer.Elapsed += (s, e) => _extEvent.Raise();
                _recoveryTimer.AutoReset = true;
                _recoveryTimer.Enabled = true;
            }

            var app = application.ControlledApplication;
            app.DocumentOpened += OnAnyDocumentOpened;
            app.DocumentClosing += OnAnyDocumentClosing;

            return Result.Succeeded;
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            try
            {
                // удалим все lock-файлы, созданные этой сессией
                RemoveAllLocksForCurrentProcess();
            }
            catch { }

            Logger.Info($"Addin {ThisVersion} shutdown");
            _watcher?.Dispose();

            if (_recoveryTimer != null)
            {
                _recoveryTimer.Stop();
                _recoveryTimer.Dispose();
            }

            application.DialogBoxShowing -= OnDialogBoxShowing;

            var app = application.ControlledApplication;
            app.DocumentOpened -= OnAnyDocumentOpened;
            app.DocumentClosing -= OnAnyDocumentClosing;

            return Result.Succeeded;
        }

        private void OnFileCreated(object sender,FileSystemEventArgs e)
        {
            if (e.FullPath.EndsWith(".json", StringComparison.OrdinalIgnoreCase) || e.FullPath.EndsWith(".running", StringComparison.OrdinalIgnoreCase))
            {
                Logger.Info("[INFO] New task in queue");
                _extEvent.Raise();
            }
            Logger.Info("[INFO] New file in queue, but not .json or .running");
        }

        // Только целевые автоклики + лог, прочие окна не трогаем
        private void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs e)
        {
            try
            {
                string id = e.DialogId ?? string.Empty;
                string msg = (e as TaskDialogShowingEventArgs)?.Message
                          ?? (e as MessageBoxShowingEventArgs)?.Message
                          ?? string.Empty;

                Logger.Info($"Dialog: Id='{id}' Msg='{msg}'");

                // 1) Duplicate name -> Overwrite existing copy (Yes)
                if (id.IndexOf("Duplicate", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Duplicate name", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Overwrite existing", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Дубликат", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Перезаписать существующую", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Перезаписать", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (e is TaskDialogShowingEventArgs td) td.OverrideResult((int)TaskDialogResult.Yes);
                    Logger.Info("  -> action: Yes (Overwrite existing copy)");
                    return;
                }

                // 2) Opening Worksets / Opening Warnings -> OK
                if (id.IndexOf("Opening Worksets", StringComparison.OrdinalIgnoreCase) >= 0
                    || id.IndexOf("Opening_Worksets", StringComparison.OrdinalIgnoreCase) >= 0
                    || id.IndexOf("OpeningWarnings", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Opening Worksets", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Открытие рабочих наборов", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Opening Warnings", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Предупрежден", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (e is TaskDialogShowingEventArgs td) td.OverrideResult((int)TaskDialogResult.Ok);
                    Logger.Info("  -> action: OK (Opening Worksets/Warnings)");
                    return;
                }

                // 3) Upgrade model prompt -> Yes
                if (msg.IndexOf("предыдущей версии", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("будет обновл", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("previous version", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("upgraded", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    if (e is TaskDialogShowingEventArgs td) td.OverrideResult((int)TaskDialogResult.Yes);
                    Logger.Info("  -> action: Yes (Upgrade model to new version)");
                    return;
                }

                Logger.Info("  -> action: none");
            }
            catch
            {
                /* не роняем Revit из-за логгера */
            }
        }

        private void OnAnyDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            try
            {
                var d = e.Document;
                var vis = GetVisiblePath(d);
                if (string.IsNullOrEmpty(vis)) return;

                var lockPath = RevitProcessLauncher.GetModelLockPath(ThisVersion, vis);
                Directory.CreateDirectory(Path.GetDirectoryName(lockPath));
                var content = _pid.ToString() + "|" + vis + "|" + DateTime.UtcNow.ToString("O");
                File.WriteAllText(lockPath, content);
                Logger.Info($"[LOCK] Created: {lockPath}");
            }
            catch (Exception ex)
            {
                Logger.Error("[LOCK] Create failed: " + ex.Message);
            }
        }

        private void OnAnyDocumentClosing(object sender, DocumentClosingEventArgs e)
        {
            try
            {
                var d = e.Document;
                var vis = GetVisiblePath(d);
                if (string.IsNullOrEmpty(vis)) return;

                var lockPath = RevitProcessLauncher.GetModelLockPath(ThisVersion, vis);
                TryRemoveLockForThisPid(lockPath);
            }
            catch (Exception ex)
            {
                Logger.Error("[LOCK] Remove on closing failed: " + ex.Message);
            }
        }

        private static string GetVisiblePath(Document d)
        {
            if (d == null) return null;
            try
            {
                if (d.IsWorkshared)
                {
                    var mp = d.GetWorksharingCentralModelPath();
                    if (mp != null)
                        return ModelPathUtils.ConvertModelPathToUserVisiblePath(mp);
                }
            }
            catch { }
            try
            {
                return d.PathName; // для не-WS моделей
            }
            catch
            {
                return null;
            }
        }

        private static void TryRemoveLockForThisPid(string lockPath)
        {
            try
            {
                if (string.IsNullOrEmpty(lockPath) || !File.Exists(lockPath)) return;

                var text = File.ReadAllText(lockPath).Trim();
                if (string.IsNullOrEmpty(text)) { File.Delete(lockPath); return; }

                var parts = text.Split('|');
                int pidInFile;
                if (parts.Length > 0 && int.TryParse(parts[0], out pidInFile))
                {
                    if (pidInFile == _pid)
                    {
                        File.Delete(lockPath);
                        Logger.Info($"[LOCK] Removed: {lockPath}");
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Error($"[LOCK] Failed to remove lock '{lockPath}': {ex.Message}");
            }
        }

        private static void RemoveAllLocksForCurrentProcess()
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AutoOpener", "locks", ThisVersion.ToString());

                if (!Directory.Exists(dir)) return;

                foreach (var lf in Directory.GetFiles(dir, "*.lock"))
                {
                    try
                    {
                        var text = File.ReadAllText(lf).Trim();
                        var parts = string.IsNullOrEmpty(text) ? Array.Empty<string>() : text.Split('|');
                        int pidInFile;
                        if (parts.Length > 0 && int.TryParse(parts[0], out pidInFile) && pidInFile == _pid)
                        {
                            File.Delete(lf);
                            Logger.Info($"[LOCK] Removed (shutdown): {lf}");
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }
    }

    /// <summary>
    /// Открытие модели на следующем тике UI: WorksetConfiguration (CloseAll + выбранные), логирование.
    /// </summary>
    public class OpenJobHandler : IExternalEventHandler
    {
        private readonly int _revitVersion;
        private bool _scheduled;
        private AutoOpenJob _pendingJob;
        private string _pendingJobFile;
        private int _currentModelIndex;
        private JobResult _jobResult;

        public OpenJobHandler(int revitVersion)
        {
            _revitVersion = revitVersion;
        }

        public bool IsScheduled => _scheduled;

        public void Execute(UIApplication uiapp)
        {
            if (_scheduled)
            {
                Logger.Info("Execute: already scheduled");
                return;
            }

            Logger.Info("Execute: start");
            var jobFile = TryTakeJobForVersion(_revitVersion);
            Logger.Info(jobFile == null ? "Execute: no job" : $"Execute: job taken {Path.GetFileName(jobFile)}");
            if (jobFile == null) return;

            try
            {
                _pendingJob = JsonStorage.Read<AutoOpenJob>(jobFile);
                Logger.Info($"Job: ver={_pendingJob.RevitVersion}, models='{_pendingJob.Models}', ws={_pendingJob.Models?.Count ?? 0}");
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to read job: " + ex);
                TryMarkBad(jobFile);
                return;
            }

            _pendingJobFile = jobFile; // файл помечен как *.running
            _currentModelIndex = 0;
            _jobResult = new JobResult
            {
                JobId = _pendingJob.Id,
                RevitVersion = _revitVersion,
                JobType = "Open",
                Succeeded = true,
                Message = "",
                OpenedModelPaths = new List<string>()
            };

            if (_pendingJob.Models == null || _pendingJob.Models.Count == 0)
            {
                _jobResult.Succeeded = false;
                _jobResult.Message = "No models to open in this job.";
                JobFiles.MarkFail(_pendingJobFile, _jobResult);
                ResetState();
                return;
            }

            _scheduled = true;
            uiapp.Idling += OnUiIdlingProcessModels;
            Logger.Info($"[STATE MACHINE] Scheduled job for {_pendingJob.Models.Count} models.");
        }

        private void OnUiIdlingProcessModels(object sender, IdlingEventArgs e)
        {
            var uiapp = sender as UIApplication;

            // Если задачи закончились или job пустой - отписываемся и формируем ответ
            if (_pendingJob == null || _currentModelIndex >= _pendingJob.Models.Count)
            {
                if (uiapp != null) uiapp.Idling -= OnUiIdlingProcessModels;
                FinalizeJob();
                return;
            }

            var currentTask = _pendingJob.Models[_currentModelIndex];
            _currentModelIndex++;

            ProcessSingleModel(uiapp, currentTask);
        }

        private void ProcessSingleModel(UIApplication uiapp, ModelTask task)
        {
            // Подпишемся на время открытия (диалоги/ошибки/документ), отпишемся в finally
            uiapp.DialogBoxShowing += OnDialogBoxShowing;
            uiapp.Application.FailuresProcessing += OnFailuresProcessing;
            uiapp.Application.DocumentOpened += OnDocumentOpened;

            string taskName = Path.GetFileName(task.ModelPath ?? "Unkown");
            Logger.Info($"[STATE MACHINE] Processing model {_currentModelIndex}/{_pendingJob.Models.Count}: {taskName}");

            try
            {
                var _modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(task.ModelPath);

                if (IsSameCentralAlreadyOpen(uiapp.Application, _modelPath))
                {
                    Logger.Info($"Model {taskName} already open.");
                    _jobResult.OpenedModelPaths.Add(task.ModelPath);
                    _jobResult.Message += $"[{taskName}] Already open. ";
                    return;
                }

                var toOpenIds = ResolveWorksetsToOpen(_modelPath, task.WorksetsByName);
                var wsc = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                if (toOpenIds.Count > 0) wsc.Open(toOpenIds.ToList());

                if (_pendingJob.CreateNewLocal)
                {
                    try
                    {
                        string docsPath = PathsService.OutDirFor(_revitVersion);
                        string originalFileName = task.ModelPath.Split(new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries).Last();
                        string nameWithoutExt = Path.GetFileNameWithoutExtension(originalFileName);
                        string ext = Path.GetExtension(originalFileName);
                        string userName = uiapp.Application.Username;

                        string localFileName = $"{nameWithoutExt}_{userName}{ext}";
                        string localPathStr = Path.Combine(docsPath, localFileName);
                        ModelPath localModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(localPathStr);

                        if (File.Exists(localPathStr)) File.Delete(localPathStr);

                        WorksharingUtils.CreateNewLocal(_modelPath, localModelPath);
                        _modelPath = localModelPath;
                        Logger.Info($"[OPEN] Created new local: {localPathStr}");

                        System.Threading.Thread.Sleep(3000);
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException) { Logger.Info("[OPEN] Model is not workshared."); }
                    catch (Exception ex) { Logger.Info($"[OPEN-WARN] Could not create local copy: {ex.Message}"); }
                }

                var openOpts = new OpenOptions
                {
                    DetachFromCentralOption = DetachFromCentralOption.DoNotDetach,
                    AllowOpeningLocalByWrongUser = true,
                    Audit = false
                };
                openOpts.SetOpenWorksetsConfiguration(wsc);

                Logger.Info($"Calling OpenAndActivateDocument for {taskName}...");
                var uiDoc = uiapp.OpenAndActivateDocument(_modelPath, openOpts, false);

                if (uiDoc != null)
                {
                    _jobResult.OpenedModelPaths.Add(uiDoc.Document.PathName);
                    _jobResult.Message += $"[{taskName}] Opened successfully. ";
                }
            }
            catch (Exception ex)
            {
                var isServer = task.ModelPath != null && task.ModelPath.StartsWith("RSN://", StringComparison.OrdinalIgnoreCase);
                var reason = ClassifyOpenError(ex, isServer);
                Logger.Error($"Failed to open {taskName}: {reason}");

                // Если модель упала, помечаем общий результат как failed, но продолжаем открывать остальные
                _jobResult.Succeeded = false;
                _jobResult.Message += $"[{taskName}] Error: {reason}. ";
            }
            finally
            {
                try
                {
                    uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                    uiapp.Application.FailuresProcessing -= OnFailuresProcessing;
                    uiapp.Application.DocumentOpened -= OnDocumentOpened;
                }
                catch (Exception ex) { Logger.Error($"[OPEN-CLEANUP] Failed to unsubscribe from events: {ex.Message}"); }
            }
        }

        private void FinalizeJob()
        {
            if (_pendingJobFile == null) return;

            Logger.Info($"[STATE MACHINE] Job finished. Overall Success: {_jobResult.Succeeded}");
            try
            {
                if (_jobResult.Succeeded)
                    JobFiles.MarkDone(_pendingJobFile, _jobResult);
                else
                    JobFiles.MarkFail(_pendingJobFile, _jobResult);
            }
            catch (Exception ex) { Logger.Error($"Failed to mark job result: {ex.Message}"); }

            ResetState();
        }

        public string GetName() => "AutoOpener OpenJob ExternalEvent";

        // ---------- локальные обработчики диалогов/ошибок/событий документа ----------
        private void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs e)
        {
            try
            {
                var id = e.DialogId ?? string.Empty;
                var msg = (e as TaskDialogShowingEventArgs)?.Message ?? (e as MessageBoxShowingEventArgs)?.Message ?? string.Empty;

                if (id.IndexOf("Duplicate", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Duplicate name", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Overwrite existing", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Дубликат", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Перезаписать", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    (e as TaskDialogShowingEventArgs)?.OverrideResult((int)TaskDialogResult.Yes);
                    return;
                }

                if (id.IndexOf("Opening Worksets", StringComparison.OrdinalIgnoreCase) >= 0 || id.IndexOf("OpeningWarnings", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Opening Worksets", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Открытие рабочих наборов", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Opening Warnings", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("Предупрежден", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    (e as TaskDialogShowingEventArgs)?.OverrideResult((int)TaskDialogResult.Ok);
                    return;
                }

                if (msg.IndexOf("предыдущей версии", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("будет обновл", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("previous version", StringComparison.OrdinalIgnoreCase) >= 0 || msg.IndexOf("upgraded", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    (e as TaskDialogShowingEventArgs)?.OverrideResult((int)TaskDialogResult.Yes);
                    return;
                }
            }
            catch { }
        }

        private void OnFailuresProcessing(object sender, FailuresProcessingEventArgs e)
        {
            try
            {
                var fa = e.GetFailuresAccessor();
                var failures = fa.GetFailureMessages();
                foreach (var f in failures)
                {
                    if (f.GetSeverity() == FailureSeverity.Warning) fa.DeleteWarning(f);
                }
            }
            catch { }
        }

        private void OnDocumentOpened(object sender, DocumentOpenedEventArgs e)
        {
            try
            {
                var d = e.Document;
                Logger.Info($"[OPEN] DocumentOpened: '{d?.PathName}' IsWorkshared={d?.IsWorkshared} Title='{d?.Title}'");
            }
            catch { }
        }

        // ---------- вспомогательные методы ----------
        private static string TryTakeJobForVersion(int version)
        {
            var dir = PathsService.QueueDirFor(version);
            if (!Directory.Exists(dir)) return null; 
            
            var files = Directory.GetFiles(dir, "*.json")
                                 .OrderBy(File.GetCreationTimeUtc)
                                 .ThenBy(File.GetLastWriteTimeUtc);

            // Блок восстановления зависших задач (Без переименования)
            var runningFiles = Directory.EnumerateFiles(dir,"*.running");
            foreach (var rf in runningFiles)
            {
                try
                {
                    var fileAgeMinutes = (DateTime.Now - File.GetLastWriteTime(rf)).TotalMinutes;
                    if (fileAgeMinutes > 10.0)
                    {
                        Logger.Info($"[RECOVERY] Resuming stuck job directly: {Path.GetFileName(rf)} (Age: {fileAgeMinutes:F1} min)");
                        // Обновляем время файла, чтобы таймер не трогал его еще 10 минут
                        File.SetLastWriteTime(rf, DateTime.Now);
                        return rf;
                    }
                    else
                    {
                        Logger.Info($"[RECOVERY-SKIP] Ignored {Path.GetFileName(rf)}. It is too fresh (Age: {fileAgeMinutes:F1} min, need > 1.0)");
                    }
                }
                catch (Exception ex) 
                { 
                    Logger.Error($"[ERROR] Failed to recover file {rf}: {ex.Message}"); 
                }
            }

            foreach (var f in files)
            {
                try
                {
                    // Сначала пробуем переименовать (атомарный захват).
                    // Если файл занят, File.Move выбросит IOException
                    var running = Path.ChangeExtension(f, ".running");
                    File.Move(f, running);

                    // Теперь читаем уже приватный файл
                    var job = JsonStorage.Read<AutoOpenJob>(running);
                    if (job == null || job.RevitVersion != version)
                    {
                        Logger.Error($"[JOB REJECTED] File '{running}' is not a valid AutoOpenJob. Moved to .bad");
                        TryMove(running, Path.ChangeExtension(running, ".bad"));
                        continue;
                    }
                    if (job.RevitVersion != version)
                    {
                        Logger.Error($"[JOB REJECTED] File '{running}' has wrong version (Expected: {version}, Got: {job.RevitVersion}). Moved to .bad"); 
                        TryMove(running, Path.ChangeExtension(running, ".bad"));
                        continue;
                    }

                    return running;
                }
                catch (IOException)
                {
                    continue;
                }
                catch (Exception ex)
                {
                    Logger.Error($"Error processing job file '{f}': {ex}");
                    TryMove(f, Path.ChangeExtension(f, ".bad"));
                }
            }
            return null;
        }

        private static bool IsSameCentralAlreadyOpen(Autodesk.Revit.ApplicationServices.Application app, ModelPath targetCentral)
        {
            var target = ModelPathUtils.ConvertModelPathToUserVisiblePath(targetCentral);
            foreach (Document d in app.Documents)
            {
                try
                {
                    if (!d.IsWorkshared) continue;
                    var central = d.GetWorksharingCentralModelPath();
                    if (central == null) continue;
                    var vis = ModelPathUtils.ConvertModelPathToUserVisiblePath(central);
                    if (StringComparer.OrdinalIgnoreCase.Equals(vis, target))
                        return true;
                }
                catch { }
            }
            return false;
        }

        private static HashSet<WorksetId> ResolveWorksetsToOpen(ModelPath mp, IList<string> names)
        {
            var result = new HashSet<WorksetId>();
            if (names == null || names.Count == 0) 
                return result;

            IList<WorksetPreview> previews = null;
            try
            {
                Logger.Info("[WORKSETS] Requesting WorksetInfo from Revit Server... (This may take 30-60 seconds, DO NOT KILL REVIT)");
                previews = WorksharingUtils.GetUserWorksetInfo(mp); // получаем список рабочих наборов (предпросмотр)
                Logger.Info($"[WORKSETS] Received {previews?.Count ?? 0} worksets from server.");
            }
            catch (Exception ex)
            {
                Logger.Error($"[WORKSETS-FAIL] Server unreachable or error: {ex.Message}");
                return result;
            }

            if (previews == null || previews.Count == 0) 
                return result;

            foreach (var name in names)
            {
                var hit = previews.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
                if (hit != null) result.Add(hit.Id);
            }
            return result;
        }

        private static void TryDelete(string path)
        {
            try 
            { 
                if (File.Exists(path)) File.Delete(path); 
            } 
            catch (Exception ex) 
            { 
                Logger.Error($"[IO] TryDelete failed for '{path}': {ex.Message}"); 
            }
        }

        private static void TryMarkBad(string path)
        {
            try 
            { 
                if (File.Exists(path)) 
                    File.Move(path, Path.ChangeExtension(path, ".bad")); 
            } 
            catch (Exception ex) 
            { 
                Logger.Error($"[IO] TryMarkBad failed for '{path}': {ex.Message}"); 
            }
        }

        private static void TryMarkFail(string path)
        {
            try 
            { 
                if (File.Exists(path)) 
                    File.Move(path, Path.ChangeExtension(path, ".fail")); 
            }
            catch (Exception ex) 
            { 
                Logger.Error($"[IO] TryMarkFail failed for '{path}': {ex.Message}"); 
            }
        }

        private static void TryMove(string from, string to)
        {
            try 
            { 
                if (File.Exists(from)) 
                    File.Move(from, to); 
            }
            catch (Exception ex) 
            { 
                Logger.Error($"[IO] TryMove failed '{from}' -> '{to}': {ex.Message}"); 
            }
        }

        private void ResetState()
        {
            _scheduled = false;
            _pendingJob = null;
            _pendingJobFile = null;
            _currentModelIndex = 0;
            _jobResult = null;
        }

        /// <summary>
        /// Приводим исключение к человеческой причине сбоя открытия.
        /// isServerPath=true → RSN-специфичное сообщение при типичных паттернах "не найдено/не доступно".
        /// </summary>
        private static string ClassifyOpenError(Exception ex, bool isServerPath)
        {
            if (ex == null) 
                return "Unknown error";

            var msg = ex.Message ?? string.Empty;
            var raw = ex.ToString() ?? string.Empty;

            // 1) Запрет смены активного документа в контексте события
            if (msg.IndexOf("Switching active documents is not allowed", StringComparison.OrdinalIgnoreCase) >= 0
                || raw.IndexOf("Switching active documents is not allowed", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Revit API prevented switching the active document during event handling. Open this model in a separate Revit session.";

            // 2) RSN-ветка: сервер/проект недоступны, не назначены или не существуют
            if (isServerPath)
            {
                if (msg.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("cannot be found", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("failed to connect", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("not available", StringComparison.OrdinalIgnoreCase) >= 0
                    || raw.IndexOf("Revit Server", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    return "Revit Server does not exist or not scheduled";
                }
            }

            // 3) Файл/путь недоступен (локальный/сетевой путь)
            if (msg.IndexOf("does not exist", StringComparison.OrdinalIgnoreCase) >= 0
                || msg.IndexOf("file not found", StringComparison.OrdinalIgnoreCase) >= 0)
                return "Model not found or unreachable (file path).";

            // 4) Общая валидация пути/аргумента
            if (ex is Autodesk.Revit.Exceptions.ArgumentException)
                return "Invalid model path or options: " + msg;

            // 5) По умолчанию — оригинальный текст
            return msg.Length > 0 ? msg : "Unknown error";
        }
    }
}
