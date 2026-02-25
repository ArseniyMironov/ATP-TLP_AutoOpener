using AutoOpener.Core.IO;
using AutoOpener.Core.Jobs;
using AutoOpener.Core.Processes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;

namespace AutoOpener.LauncherCLI
{
    class Program
    {
        static int Main(string[] args)
        {
            // Пример: AutoOpener.LauncherCLI.exe enqueue 2022 "RSN://srv/.../model.rvt" "AR_Walls;AR_Stairs"
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: enqueue <version> <rsnPath> [worksets(;)] [--wait]");
                return 1;
            }

            var cmd = args[0];
            var version = int.Parse(args[1]);
            var rsn = args[2];

            PathsService.SetVersionContext(version);

            bool wait = false;
            var worksets = new List<string>();

            // Парсим дополнительные аргументы (рабочие наборы и флаг ожидания)
            for (int i = 3; i < args.Length; i++)
            {
                if (args[i].Equals("--wait", StringComparison.OrdinalIgnoreCase))
                    wait = true;
                else
                    worksets.AddRange(args[i].Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries));
            }

            var job = new AutoOpenJob
            {
                Id = Guid.NewGuid(),
                RevitVersion = version,
                RsnPath = rsn,
                WorksetsByName = worksets,
                CreateNewLocal = true
            };

            var jobFile = Path.Combine(PathsService.QueueDirFor(version), job.Id + ".json");
            JsonStorage.Write(jobFile, job);
            CleanupService.CleanupOldArtifacts(version, 7);

            if (cmd.Equals("open-now", StringComparison.OrdinalIgnoreCase) ||
                cmd.Equals("enqueue", StringComparison.OrdinalIgnoreCase))
            {
                if (!RevitProcessLauncher.TryStart(version, rsn))
                {
                    Console.WriteLine("Revit.exe not found for version " + version);
                    return 2;
                }

                if (!wait)
                {
                    Console.WriteLine($"Job {job.Id} queued and Revit started.");
                    return 0;
                }

                // Логика ожидания результата (--wait)
                Console.WriteLine($"Waiting for result of job {job.Id}...");
                var resultFile = JobFiles.GetResultPath(job.Id);
                int timeoutSeconds = 600;
                int elapsed = 0;

                while (elapsed < timeoutSeconds)
                {
                    if (File.Exists(resultFile))
                    {
                        Thread.Sleep(500);  // Даем Revit время закрыть файл (File Handle)
                        try
                        {
                            var res = JsonStorage.Read<JobResult>(resultFile);
                            if (res.Succeeded)
                            {
                                Console.WriteLine($"\n[SUCCESS] {res.Message}");
                                Console.WriteLine($"Local path: {res.ModelPath}");
                                return 0; // Exit Code 0 = Успех
                            }
                            else
                            {
                                Console.WriteLine($"\n[ERROR] {res.Message}");
                                return 1; // Exit Code 1 = Ошибка при открытии
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"\n[ERROR] Failed to read result file: {ex.Message}");
                            return 1;
                        }
                    }

                    Thread.Sleep(2000);  // Проверяем каждые 2 секунды
                    elapsed += 2;
                    Console.Write('.');  // Выводим точки в консоль для понимания, что процесс жив
                }

                Console.WriteLine("\n[TIMEOUT] Revit did not process the job in time (10 minutes).");
                return 3; // Exit Code 3 = Таймаут
            }

            Console.WriteLine("Unkown command.");
            return 1;
        }
    }
}
