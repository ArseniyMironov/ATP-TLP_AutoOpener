using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Events;
using AutoOpener.Core.IO;
using AutoOpener.Core.Jobs;
using AutoOpener.Core.Processes;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace AutoOpener.Addin2023
{
    public class App : IExternalApplication
    {
        private const int ThisVersion = 2023;

        private ExternalEvent _extEvent;
        private OpenJobHandler _handler;
        private FileSystemWatcher _watcher;
        private static readonly int _pid = Process.GetCurrentProcess().Id;

        public Result OnStartup(UIControlledApplication application)
        {
            PathsService.SetVersionContext(ThisVersion);
            Directory.CreateDirectory(PathsService.LogsDir);
            Directory.CreateDirectory(PathsService.OutDir);
            CleanupService.CleanupOldArtifacts(ThisVersion, 7);
            Logger.Info("Addin 2023 startup");

            _handler = new OpenJobHandler(ThisVersion);
            _extEvent = ExternalEvent.Create(_handler);

            application.DialogBoxShowing += OnDialogBoxShowing;

            try
            {
                _watcher = new FileSystemWatcher(PathsService.QueueDirFor(ThisVersion));
                _watcher.Filter = "*.json";
                _watcher.Created += OnFileCreated;
                _watcher.EnableRaisingEvents = true;

                _extEvent.Raise();
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to start FileSystemWatcherL " + ex);
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

            Logger.Info("Addin 2022 shutdown");
            _watcher?.Dispose();
            application.DialogBoxShowing -= OnDialogBoxShowing;

            var app = application.ControlledApplication;
            app.DocumentOpened -= OnAnyDocumentOpened;
            app.DocumentClosing -= OnAnyDocumentClosing;

            return Result.Succeeded;
        }

        private void OnFileCreated(object sender, FileSystemEventArgs e)
        {
            _extEvent.Raise();
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
            catch { /* не роняем Revit из-за lock-файла */ }
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
        private ModelPath _modelPath;

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

            AutoOpenJob job;
            try
            {
                job = JsonStorage.Read<AutoOpenJob>(jobFile);
                Logger.Info($"Job: ver={job.RevitVersion}, path='{job.RsnPath}', ws={job.WorksetsByName?.Count ?? 0}");
            }
            catch (Exception ex)
            {
                Logger.Error("Failed to read job: " + ex);
                TryMarkBad(jobFile);
                return;
            }

            _pendingJobFile = jobFile; // файл помечен как *.running
            _pendingJob = job;

            try
            {
                _modelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(_pendingJob.RsnPath);
                Logger.Info("ModelPath converted OK");
            }
            catch (Exception ex)
            {
                Logger.Error("Invalid RSN/path: " + ex);
                try
                {
                    var res = new JobResult
                    {
                        JobId = _pendingJob != null ? _pendingJob.Id : Guid.Empty,
                        Succeeded = false,
                        Message = "Invalid RSN/path: " + ex.Message,
                        JobType = "Open",
                        ModelPath = _pendingJob != null ? _pendingJob.RsnPath : null
                    };
                    JobFiles.MarkFail(_pendingJobFile ?? "", res);
                }
                catch { }
                ResetState();
                return;
            }

            if (IsSameCentralAlreadyOpen(uiapp.Application, _modelPath))
            {
                Logger.Info("Model already open → mark done");
                try
                {
                    var res = new JobResult
                    {
                        JobId = _pendingJob.Id,
                        Succeeded = true,
                        Message = "Model already open",
                        JobType = "Open",
                        ModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(_modelPath)
                    };
                    JobFiles.MarkDone(_pendingJobFile, res);
                }
                catch { }
                ResetState();
                return;
            }

            _scheduled = true;
            uiapp.Idling += OnUiIdlingOpenOnce;
            Logger.Info("Scheduled open on next Idling");
        }

        private void OnUiIdlingOpenOnce(object sender, IdlingEventArgs e)
        {
            var uiapp = sender as UIApplication;
            if (uiapp != null) uiapp.Idling -= OnUiIdlingOpenOnce;

            if (_pendingJob == null || _modelPath == null)
            {
                Logger.Info("IdlingOpen: no pending job/state");
                ResetState();
                return;
            }

            // Подпишемся на время открытия (диалоги/ошибки/документ), отпишемся в finally
            uiapp.DialogBoxShowing += OnDialogBoxShowing;
            uiapp.Application.FailuresProcessing += OnFailuresProcessing;
            uiapp.Application.DocumentOpened += OnDocumentOpened;

            try
            {
                Logger.Info("IdlingOpen: resolve worksets to open");
                var toOpenIds = ResolveWorksetsToOpen(_modelPath, _pendingJob.WorksetsByName);
                Logger.Info($"IdlingOpen: toOpen={toOpenIds.Count}");

                // вычислим отсутствующие WS (для сообщения)
                var missingWs = new List<string>();
                try
                {
                    if (_pendingJob.WorksetsByName != null && _pendingJob.WorksetsByName.Count > 0)
                    {
                        IList<WorksetPreview> previews = null;
                        try { previews = WorksharingUtils.GetUserWorksetInfo(_modelPath); } catch { previews = null; }
                        var names = previews != null
                            ? new HashSet<string>(previews.Select(p => p.Name), StringComparer.OrdinalIgnoreCase)
                            : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        foreach (var n in _pendingJob.WorksetsByName)
                            if (!names.Contains(n)) missingWs.Add(n);
                    }
                }
                catch { /* диагностическое, не критично */ }

                if (missingWs.Count > 0)
                    Logger.Info("[OPEN] Missing worksets: " + string.Join("; ", missingWs));

                var wsc = new WorksetConfiguration(WorksetConfigurationOption.CloseAllWorksets);
                if (toOpenIds.Count > 0) wsc.Open(toOpenIds.ToList());

                if (_pendingJob.CreateNewLocal)
                {
                    try
                    {
                        // 1. Формируем путь для локального файла (%AppData%\AutoOpener\{version}\out)
                        string docsPath = PathsService.OutDirFor(_revitVersion);
                        string fileName = GetUniqueLocalFileName(_pendingJob.RsnPath);
                        string localPathStr = Path.Combine(docsPath, fileName);
                        ModelPath localModelPath = ModelPathUtils.ConvertUserVisiblePathToModelPath(localPathStr);

                        // 2. Пытаемся удалить старый файл (строгая перезапись)
                        if (File.Exists(localPathStr))
                        {
                            // Если файл занят, тут вылетит IOException -> задача упадет (как и требовалось)
                            File.Delete(localPathStr);
                        }

                        // 3. Пытаемся создать локальную копию
                        // Если модель НЕ workshared, Revit API выбросит ArgumentException
                        WorksharingUtils.CreateNewLocal(_modelPath, localModelPath);

                        // 4. Успех -> подменяем путь на локальный
                        _modelPath = localModelPath;
                        Logger.Info($"[OPEN] Created new local: {localPathStr}");
                    }
                    catch (Autodesk.Revit.Exceptions.ArgumentException)
                    {
                        // Исключение: "Model is not workshared"
                        // Согласно логике: "если не worksharing - открываем его (оригинал)"
                        Logger.Info("[OPEN] Model is not workshared. Opening original path.");
                    }
                    catch (Exception ex)
                    {
                        // Если модель не workshared или другая ошибка — просто пишем лог и открываем оригинал
                        Logger.Info($"[OPEN-WARN] Could not create local copy (opening original): {ex.Message}");
                    }
                }

                var openOpts = new OpenOptions
                {
                    DetachFromCentralOption = DetachFromCentralOption.DoNotDetach,
                    AllowOpeningLocalByWrongUser = true,
                    Audit = false
                };
                openOpts.SetOpenWorksetsConfiguration(wsc);

                Logger.Info("IdlingOpen: calling OpenDocumentFile...");
                var uiDoc = uiapp.OpenAndActivateDocument(_modelPath, openOpts, false);
                Logger.Info($"IdlingOpen: OpenDocumentFile returned, doc='{uiDoc?.Document.Title}'");

                // формируем сообщение об успехе
                string msg = "Opened successfully";
                if (missingWs.Count > 0)
                    msg += ". Missing worksets: " + string.Join("; ", missingWs);

                try
                {
                    var res = new JobResult
                    {
                        JobId = _pendingJob.Id,
                        Succeeded = true,
                        Message = msg,
                        JobType = "Open",
                        ModelPath = ModelPathUtils.ConvertModelPathToUserVisiblePath(_modelPath)
                    };
                    JobFiles.MarkDone(_pendingJobFile, res);
                }
                catch { }

                Logger.Info("IdlingOpen: job marked done, finished.");
            }
            catch (Exception ex)
            {
                var isServer = !string.IsNullOrEmpty(_pendingJob?.RsnPath)
                    && _pendingJob.RsnPath.TrimStart().StartsWith("RSN://", StringComparison.OrdinalIgnoreCase);
                var reason = ClassifyOpenError(ex, isServer);

                Logger.Error("IdlingOpen failed: " + reason + " | Raw: " + ex);

                try
                {
                    var res = new JobResult
                    {
                        JobId = _pendingJob != null ? _pendingJob.Id : Guid.Empty,
                        Succeeded = false,
                        Message = reason,
                        JobType = "Open",
                        ModelPath = _pendingJob != null ? _pendingJob.RsnPath : null
                    };
                    JobFiles.MarkFail(_pendingJobFile ?? "", res);
                }
                catch { }

                try
                {
                    TaskDialog.Show("AutoOpener 2022", "Open failed:\n" + reason);
                }
                catch { }
            }
            finally
            {
                try
                {
                    uiapp.DialogBoxShowing -= OnDialogBoxShowing;
                    uiapp.Application.FailuresProcessing -= OnFailuresProcessing;
                    uiapp.Application.DocumentOpened -= OnDocumentOpened;
                }
                catch { }
                ResetState();
            }
        }

        public string GetName() => "AutoOpener OpenJob ExternalEvent";

        // ---------- локальные обработчики диалогов/ошибок/событий документа ----------
        private void OnDialogBoxShowing(object sender, DialogBoxShowingEventArgs e)
        {
            try
            {
                var id = e.DialogId ?? string.Empty;
                var msg = (e as TaskDialogShowingEventArgs)?.Message
                       ?? (e as MessageBoxShowingEventArgs)?.Message
                       ?? string.Empty;

                Logger.Info($"[OPEN] Dialog: Id='{id}' Msg='{msg}'");

                // Duplicate name → Yes (Overwrite)
                if (id.IndexOf("Duplicate", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Duplicate name", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Overwrite existing", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Дубликат", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Перезаписать", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    (e as TaskDialogShowingEventArgs)?.OverrideResult((int)TaskDialogResult.Yes);
                    Logger.Info("  -> action: Yes (overwrite)");
                    return;
                }

                // Opening Worksets/Warnings → OK
                if (id.IndexOf("Opening Worksets", StringComparison.OrdinalIgnoreCase) >= 0
                    || id.IndexOf("Opening_Worksets", StringComparison.OrdinalIgnoreCase) >= 0
                    || id.IndexOf("OpeningWarnings", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Opening Worksets", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Открытие рабочих наборов", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Opening Warnings", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("Предупрежден", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    (e as TaskDialogShowingEventArgs)?.OverrideResult((int)TaskDialogResult.Ok);
                    Logger.Info("  -> action: OK (opening worksets/warnings)");
                    return;
                }

                // Older version model prompt → Yes (Upgrade)
                if (msg.IndexOf("предыдущей версии", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("будет обновл", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("previous version", StringComparison.OrdinalIgnoreCase) >= 0
                    || msg.IndexOf("upgraded", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    (e as TaskDialogShowingEventArgs)?.OverrideResult((int)TaskDialogResult.Yes);
                    Logger.Info("  -> action: Yes (upgrade model)");
                    return;
                }

                Logger.Info("  -> action: none");
            }
            catch { }
        }

        private void OnFailuresProcessing(object sender, FailuresProcessingEventArgs e)
        {
            try
            {
                var fa = e.GetFailuresAccessor();
                var failures = fa.GetFailureMessages();
                int warnings = 0, others = 0;

                foreach (var f in failures)
                {
                    if (f.GetSeverity() == FailureSeverity.Warning)
                    {
                        fa.DeleteWarning(f);
                        warnings++;
                    }
                    else
                    {
                        others++;
                    }
                }

                if (warnings > 0 || others > 0)
                    Logger.Info($"[OPEN] Failures: deleted warnings={warnings}, other={others}");
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
        private static string GetUniqueLocalFileName(string rsnPath)
        {
            if (string.IsNullOrEmpty(rsnPath)) return "model.rvt";

            var name = Path.GetFileNameWithoutExtension(rsnPath);
            var ext = Path.GetExtension(rsnPath);

            using (var md5 = MD5.Create())
            {
                var inputBytes = Encoding.UTF8.GetBytes(rsnPath.ToLowerInvariant());
                var hashBytes = md5.ComputeHash(inputBytes);
                var sb = new StringBuilder();
                for (int i = 0; i < hashBytes.Length; i++)
                    sb.Append(hashBytes[i].ToString("X2"));

                // Берем первые 8 символов хеша для уникальности
                return $"{name}_{sb.ToString().Substring(0, 8)}{ext}";
            }
        }

        private static string TryTakeJobForVersion(int version)
        {
            var dir = PathsService.QueueDirFor(version);
            if (!Directory.Exists(dir)) return null;

            var files = Directory.GetFiles(dir, "*.json")
                                 .OrderBy(File.GetCreationTimeUtc)
                                 .ThenBy(File.GetLastWriteTimeUtc);

            foreach (var f in files)
            {
                try
                {
                    // Сначала пробуем переименовать (атомарный захват).
                    // Если файл занят (пишется), File.Move выбросит IOException
                    var running = Path.ChangeExtension(f, ".running");
                    File.Move(f, running);

                    // Теперь читаем уже наш приватный файл
                    var job = JsonStorage.Read<AutoOpenJob>(running);
                    if (job == null || job.RevitVersion != version)
                    {
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
            if (names == null || names.Count == 0) return result;

            IList<WorksetPreview> previews = null;
            try
            {
                previews = WorksharingUtils.GetUserWorksetInfo(mp); // получаем список рабочих наборов (предпросмотр)
            }
            catch
            {
                return result;
            }
            if (previews == null || previews.Count == 0) return result;

            foreach (var name in names)
            {
                var hit = previews.FirstOrDefault(p => string.Equals(p.Name, name, StringComparison.OrdinalIgnoreCase));
                if (hit != null) result.Add(hit.Id);
            }
            return result;
        }

        private static void TryDelete(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); } catch { }
        }

        private static void TryMarkBad(string path)
        {
            try { if (File.Exists(path)) File.Move(path, Path.ChangeExtension(path, ".bad")); } catch { }
        }

        private static void TryMarkFail(string path)
        {
            try { if (File.Exists(path)) File.Move(path, Path.ChangeExtension(path, ".fail")); } catch { }
        }

        private static void TryMove(string from, string to)
        {
            try { if (File.Exists(from)) File.Move(from, to); } catch { }
        }

        private void ResetState()
        {
            _scheduled = false;
            _pendingJob = null;
            _pendingJobFile = null;
            _modelPath = null;
        }

        /// <summary>
        /// Приводим исключение к человеческой причине сбоя открытия.
        /// isServerPath=true → RSN-специфичное сообщение при типичных паттернах "не найдено/не доступно".
        /// </summary>
        private static string ClassifyOpenError(Exception ex, bool isServerPath)
        {
            if (ex == null) return "Unknown error";

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
