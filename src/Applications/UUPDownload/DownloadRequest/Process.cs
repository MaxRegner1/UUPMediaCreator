/*
 * Copyright (c) Gustave Monce and Contributors
 * 
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 * 
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 * 
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;
using System.Threading.Tasks;
using UnifiedUpdatePlatform.Services.Composition.Database;
using UnifiedUpdatePlatform.Services.WindowsUpdate;
using UUPDownload.Downloading;

namespace UUPDownload.DownloadRequest
{
    public static class Process
    {
        internal static async Task ParseDownloadOptionsAsync(DownloadRequestOptions opts)
        {
            await CheckAndDownloadUpdatesAsync(
                opts.ReportingSku,
                opts.ReportingVersion,
                opts.MachineType,
                opts.FlightRing,
                opts.FlightingBranchName,
                opts.BranchReadinessLevel,
                opts.CurrentBranch,
                opts.ReleaseType,
                opts.SyncCurrentVersionOnly,
                opts.ContentType,
                opts.Mail,
                opts.Password,
                opts.OutputFolder,
                opts.Language,
                opts.Edition);
        }

        internal static async Task ParseReplayOptionsAsync(DownloadReplayOptions opts)
        {
            var update = JsonSerializer.Deserialize<UpdateData>(await File.ReadAllTextAsync(opts.ReplayMetadata));

            Logging.Log("Title: " + update.Xml.LocalizedProperties.Title);
            Logging.Log("Description: " + update.Xml.LocalizedProperties.Description);

            if (opts.Fixup.HasValue)
            {
                await ApplyFixUpAsync(update, opts.Fixup.Value, opts.AppxRoot, opts.CabsRoot);
            }
            else
            {
                await ProcessUpdateAsync(update, opts.OutputFolder, opts.MachineType, opts.Language, opts.Edition, true);
            }
        }

        private static async Task ApplyFixUpAsync(UpdateData update, Fixup specifiedFixup, string appxRoot, string cabsRoot)
        {
            switch (specifiedFixup)
            {
                case Fixup.Appx:
                    await ApplyAppxFixUpAsync(update, appxRoot, cabsRoot);
                    break;
            }
        }

        private static async Task ApplyAppxFixUpAsync(UpdateData update, string appxRoot, string cabsRoot)
        {
            if (string.IsNullOrWhiteSpace(cabsRoot))
            {
                cabsRoot = appxRoot;
            }

            Logging.Log("Building appx license map from cabs...");
            var appxLicenseFileMap = FeatureManifestService.GetAppxPackageLicenseFileMapFromCabs(Directory.GetFiles(cabsRoot, "*.cab", SearchOption.AllDirectories));
            var appxFiles = Directory.GetFiles(Path.GetFullPath(appxRoot), "appx_*", SearchOption.TopDirectoryOnly);

            if (!update.CompDBs.Any(db => db.AppX != null))
            {
                Logging.Log("Current replay is missing some appx package metadata. Re-downloading compdbs...");
                update.CompDBs = await update.GetCompDBsAsync();
            }

            var canonicalCompdb = update.CompDBs
                .Where(compDB => compDB.Tags.Tag.Any(x => x.Name.Equals("UpdateType", StringComparison.InvariantCultureIgnoreCase) && x.Value?.Equals("Canonical", StringComparison.InvariantCultureIgnoreCase) == true))
                .FirstOrDefault(x => x.AppX != null);

            if (canonicalCompdb != null)
            {
                foreach (var appxFile in appxFiles)
                {
                    string payloadHash;
                    using (var fileStream = File.OpenRead(appxFile))
                    {
                        using var sha = SHA256.Create();
                        payloadHash = Convert.ToBase64String(sha.ComputeHash(fileStream));
                    }

                    var package = canonicalCompdb.AppX.AppXPackages.Package.FirstOrDefault(p => p.Payload.PayloadItem.FirstOrDefault()?.PayloadHash == payloadHash);
                    if (package == null)
                    {
                        Logging.Log($"Could not locate package with payload hash {payloadHash}. Skipping.");
                    }
                    else
                    {
                        var appxFolder = Path.Combine(appxRoot, Path.GetDirectoryName(package.Payload.PayloadItem.FirstOrDefault()?.Path));
                        if (!Directory.Exists(appxFolder))
                        {
                            Logging.Log($"Creating {appxFolder}");
                            Directory.CreateDirectory(appxFolder);
                        }

                        var appxPath = Path.Combine(appxRoot, package.Payload.PayloadItem.FirstOrDefault()?.Path);
                        Logging.Log($"Moving {appxFile} to {appxPath}");
                        File.Move(appxFile, appxPath, true);
                    }
                }

                foreach (var package in canonicalCompdb.AppX.AppXPackages.Package)
                {
                    if (package.LicenseData != null)
                    {
                        var appxFolder = Path.Combine(appxRoot, Path.GetDirectoryName(package.Payload.PayloadItem.FirstOrDefault()?.Path));
                        if (!Directory.Exists(appxFolder))
                        {
                            Logging.Log($"Creating {appxFolder}");
                            Directory.CreateDirectory(appxFolder);
                        }

                        var appxPath = Path.Combine(appxRoot, package.Payload.PayloadItem.FirstOrDefault()?.Path);
                        var appxLicensePath = Path.Combine(appxFolder, appxLicenseFileMap[Path.GetFileName(appxPath)]);
                        Logging.Log($"Writing license to {appxLicensePath}");
                        await File.WriteAllTextAsync(appxLicensePath, package.LicenseData);
                    }
                }
            }

            Logging.Log("Appx fixup applied.");
        }

        private static async Task CheckAndDownloadUpdatesAsync(
            OSSkuId reportingSku,
            string reportingVersion,
            MachineType machineType,
            string flightRing,
            string flightingBranchName,
            string branchReadinessLevel,
            string currentBranch,
            string releaseType,
            bool syncCurrentVersionOnly,
            string contentType,
            string mail,
            string password,
            string outputFolder,
            string language,
            string edition)
        {
            Logging.Log("Checking for updates...");

            var ctac = new CTAC(reportingSku, reportingVersion, machineType, flightRing, flightingBranchName, branchReadinessLevel, currentBranch, releaseType, syncCurrentVersionOnly, contentType);
            string token = string.Empty;
            if (!string.IsNullOrEmpty(mail) && !string.IsNullOrEmpty(password))
            {
                token = await MBIHelper.GenerateMicrosoftAccountTokenAsync(mail, password);
            }

            var data = await FE3Handler.GetUpdatesAsync(null, ctac, token, FileExchangeV3UpdateFilter.ProductRelease);

            if (!data.Any())
            {
                Logging.Log("No updates found that matched the specified criteria.", Logging.LoggingLevel.Error);
            }
            else
            {
                Logging.Log($"Found {data.Count()} update(s):");

                foreach (var (update, i) in data.Select((update, i) => (update, i)))
                {
                    Logging.Log($"{i}: Title: {update.Xml.LocalizedProperties.Title}");
                    Logging.Log($"{i}: Description: {update.Xml.LocalizedProperties.Description}");
                }

                foreach (var update in data)
                {
                    Logging.Log("Title: " + update.Xml.LocalizedProperties.Title);
                    Logging.Log("Description: " + update.Xml.LocalizedProperties.Description);

                    await ProcessUpdateAsync(update, outputFolder, machineType, language, edition, true);
                }
            }
            Logging.Log("Completed.");
            if (Debugger.IsAttached)
            {
                _ = Console.ReadLine();
            }
        }

        private static async Task ProcessUpdateAsync(UpdateData update, string outputFolder, MachineType machineType, string language = "", string edition = "", bool writeMetadata = true)
        {
            string buildstr = "";
            IEnumerable<string> languages = null;

            Logging.Log("Gathering update metadata...");

            var compDBs = await update.GetCompDBsAsync();

            await Task.WhenAll(
                Task.Run(async () => buildstr = await update.GetBuildStringAsync()),
                Task.Run(async () => languages = await update.GetAvailableLanguagesAsync()));

            buildstr ??= "";

            if (buildstr.Contains("GitEnlistment(winpbld)"))
            {
                CompDB selectedCompDB = null;
                Version currentHighest = null;
                foreach (var compDB in compDBs)
                {
                    if (compDB.TargetOSVersion != null && Version.TryParse(compDB.TargetOSVersion, out var currentVer))
                    {
                        if (currentHighest == null || currentVer > currentHighest)
                        {
                            if (!string.IsNullOrEmpty(compDB.TargetBuildInfo) && !string.IsNullOrEmpty(compDB.TargetOSVersion))
                            {
                                currentHighest = currentVer;
                                selectedCompDB = compDB;
                            }
                        }
                    }
                }

                if (selectedCompDB != null)
                {
                    buildstr = $"{selectedCompDB.TargetOSVersion} ({selectedCompDB.TargetBuildInfo.Split(".")[0]}.{selectedCompDB.TargetBuildInfo.Split(".")[3]})";
                }
            }

            if (string.IsNullOrEmpty(buildstr) && update.Xml.LocalizedProperties.Title.Contains("(UUP-CTv2)"))
            {
                var unformattedBase = update.Xml.LocalizedProperties.Title.Split(" ")[0];
                buildstr = $"10.0.{unformattedBase.Split(".")[0]}.{unformattedBase.Split(".")[1]} ({unformattedBase.Split(".")[2]}.{unformattedBase.Split(".")[3]})";
            }
            else if (string.IsNullOrEmpty(buildstr))
            {
                buildstr = update.Xml.LocalizedProperties.Title;
            }

            Logging.Log("Build String: " + buildstr);
            Logging.Log("Languages: " + string.Join(", ", languages));

            await UnifiedUpdatePlatform.Services.WindowsUpdate.Downloads.UpdateUtils.ProcessUpdateAsync(update, outputFolder, machineType, new ReportProgress(), language, edition, writeMetadata);
        }
    }
}
