using AutoOpener.Core.IO;
using AutoOpener.Core.Jobs;
using AutoOpener.Core.Models;
using AutoOpener.Core.Processes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace AutoOpener.LauncherCLI
{
    class Program
    {
        static int Main(string[] args)
        {
            // Пример: AutoOpener.LauncherCLI.exe enqueue 2022 "RSN://srv/.../model.rvt" "AR_Walls;AR_Stairs"
            if (args.Length < 2)
            {
                Console.WriteLine("Usage:");
                Console.WriteLine("  enqueue <version> <rsnPath> [worksets(;)] [--wait]");
                Console.WriteLine("  run-profile <ProfileName> [--wait]");
                return 1;
            }

            var cmd = args[0];

            bool wait = args.Any(a => a.Equals("--wait", StringComparison.OrdinalIgnoreCase));

            AutoOpenJob job = null;
            int version = 0;

            if (cmd.Equals("run-profile", StringComparison.OrdinalIgnoreCase))
            {
                string profileName = args[1];
                string profilePath = Path.Combine(PathsService.ProfileDir, profileName + ".json");

                if (!File.Exists(profilePath))
                {
                    Console.WriteLine($"Profile not found: {profilePath}");
                    return 1;
                }

                var profile = JsonStorage.Read<Profile>(profilePath);
                if (profile == null || profile.Models == null || profile.Models.Count == 0)
                {
                    Console.WriteLine($"Profile is invalid or contains no models");
                    return 1;
                }

                version = profile.RevitVersion;
                PathsService.SetVersionContext(version);

                job = new AutoOpenJob
                {
                    Id = Guid.NewGuid(),
                    RevitVersion = version,
                    CreateNewLocal = true,
                    Models = profile.Models
                };
            }
            else if (cmd.Equals("enqueue", StringComparison.OrdinalIgnoreCase))
            {
                if (args.Length < 3) return 1;
                version = int.Parse(args[1]);
                string rsn = args[2];
                PathsService.SetVersionContext(version);

                var worksets = new List<string>();
                for (int i = 3; i <args.Length; i++)
                {
                    if (!args[3].StartsWith("--"))
                        worksets.AddRange(args[i].Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries));
                }

                job = new AutoOpenJob
                {
                    Id = Guid.NewGuid(),
                    RevitVersion = version,
                    CreateNewLocal = true,
                    Models = new List<ModelTask> { new ModelTask { ModelPath = rsn, WorksetsByName = worksets } }
                };
            }
            else
            {
                Console.WriteLine("Unknown command.");
                return 1;
            }

            var jobFile = Path.Combine(PathsService.QueueDirFor(version), job.Id + ".json");
            JsonStorage.Write(jobFile, job);
            CleanupService.CleanupOldArtifacts(version, 7);

            string firstModelpath = job.Models.FirstOrDefault()?.ModelPath;
            if (!RevitProcessLauncher.TryStart(version, firstModelpath))
            {
                Console.WriteLine("Revit.exe not found for version " + version);
                return 2;
            }

            if (!wait)
            {
                Console.WriteLine($"Job {job.Id} queued and Revit started.");
                return 0;
            }

            Console.WriteLine($"Waiting for result of job {job.Id}...");
            var resultFile = JobFiles.GetResultPath(job.Id);
            int timeoutSeconds = 600;
            int elapsed = 0;
            
            while (elapsed < timeoutSeconds)
            {
                if (File.Exists(resultFile))
                {
                    Thread.Sleep(500);
                    try
                    {
                        var res = JsonStorage.Read<JobResult>(resultFile);
                        if (res != null)
                        {
                            if (res.Succeeded)
                            {
                                Console.WriteLine($"\n[SUCCESS] {res.Message}");
                                Console.WriteLine($"Local paths: {string.Join(", ", res.OpenedModelPaths)}");
                                return 0;
                            }
                            else
                            {
                                Console.WriteLine($"\n[ERROR] {res.Message}");
                                return 1;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"\n[ERROR] Failed to read result file: {ex.Message}");
                        return 1;
                    }
                }

                Thread.Sleep(2000);
                elapsed += 2;
                Console.Write(".");
            }

            Console.WriteLine("\n[TIMEOUT] Revit did not process the job in time (10 minutes).");
            return 3;
        }
    }
}
