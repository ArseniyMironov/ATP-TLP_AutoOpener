using AutoOpener.Core.IO;
using AutoOpener.Core.Jobs;
using AutoOpener.Core.Processes;
using System;
using System.Collections.Generic;
using System.IO;

namespace AutoOpener.LauncherCLI
{
    class Program
    {
        static int Main(string[] args)
        {
            // Пример: AutoOpener.LauncherCLI.exe enqueue 2022 "RSN://srv/.../model.rvt" "AR_Walls;AR_Stairs"
            if (args.Length < 3)
            {
                Console.WriteLine("Usage: enqueue <version> <rsnPath> <worksets(;)> | open-now ...");
                return 1;
            }

            var cmd = args[0];
            var version = int.Parse(args[1]);
            var rsn = args[2];
            var ws = args.Length > 3 ? args[3].Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries) : new string[0];

            var job = new AutoOpenJob
            {
                Id = Guid.NewGuid(),
                RevitVersion = version,
                RsnPath = rsn,
                WorksetsByName = new List<string>(ws),
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
                Console.WriteLine("Job queued and Revit started.");
                return 0;
            }

            Console.WriteLine("Unkown command.");
            return 1;
        }
    }
}
