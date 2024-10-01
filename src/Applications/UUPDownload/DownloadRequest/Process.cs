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
using System.Data;
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
        internal static void ParseDownloadOptions(DownloadRequestOptions opts)
        {
            // Handling CoreOS 11 and Pre-development builds
            if (opts.IsCoreOS || opts.IsPreDev)
            {
                CheckAndDownloadCoreOSPreDevUpdates(
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
                    opts.Edition).Wait();
            }
            else
            {
                CheckAndDownloadUpdates(
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
                    opts.Edition).Wait();
            }
        }

        internal static void ParseReplayOptions(DownloadReplayOptions opts)
        {
            UpdateData update = JsonSerializer.Deserialize<UpdateData>(File.ReadAllText(opts.ReplayMetadata));

            Logging.Log("Title: " + update.Xml.LocalizedProperties.Title);
            Logging.Log("Description: " + update.Xml.LocalizedProperties.Description);

            if (opts.Fixup.HasValue)
            {
                ApplyFixUpAsync(update, opts.Fixup.Value, opts.AppxRoot, opts.CabsRoot).Wait();
            }
            else
            {
                ProcessUpdateAsync(update, opts.OutputFolder, opts.MachineType, opts.Language, opts.Edition, true).Wait();
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

            Logging.Log($"Building appx license map from cabs...");
            IDictionary<string, string> appxLicenseFileMap = FeatureManifestService.GetAppxPackageLicenseFileMapFromCabs(Directory.GetFiles(cabsRoot, "*.cab", SearchOption.AllDirectories));
            string[] appxFiles = Directory.GetFiles(Path.GetFullPath(appxRoot), "appx_*", SearchOption.TopDirectoryOnly);

            if (!update.CompDBs.Any(db => db.AppX != null))
            {
                Logging.Log($"Current replay is missing some appx package metadata. Re-downloading compdbs...");
                update.CompDBs = await update.GetCompDBsAsync();
            }

            CompDB canonicalCompdb = update.CompDBs
                .Where(compDB => compDB.Tags.Tag
                .Find(x => x.Name
                .Equals("UpdateType", StringComparison.InvariantCultureIgnoreCase))?.Value?
                .Equals("Canonical", StringComparison.InvariantCultureIgnoreCase) == true)
                .Where(x => x.AppX != null)
                .FirstOrDefault();

            if (canonicalCompdb != null)
            {
                foreach (string appxFile in appxFiles)
                {
                    string payloadHash;
                    using (FileStream fileStream = File.OpenRead(appxFile))
                    {
                        using SHA256 sha = SHA256.Create();
                        payloadHash = Convert.ToBase64String(sha.ComputeHash(fileStream));
                    }

                    AppxPackage package = canonicalCompdb.AppX.AppXPackages.Package.Where(p => p.Payload.PayloadItem.FirstOrDefault().PayloadHash == payloadHash).FirstOrDefault();
                    if (package == null)
                    {
                        Logging.Log($"Could not locate package with payload hash {payloadHash}. Skipping.");
                    }
                    else
                    {
                        string appxFolder = Path.Combine(appxRoot, Path.GetDirectoryName(package.Payload.PayloadItem.FirstOrDefault().Path));
                        if (!Directory.Exists(appxFolder))
                        {
                            Logging.Log($"Creating {appxFolder}");
                            _ = Directory.CreateDirectory(appxFolder);
                        }

                        string appxPath = Path.Combine(appxRoot, package.Payload.PayloadItem.FirstOrDefault().Path);
                        Logging.Log($"Moving {appxFile} to {appxPath}");
                        File.Move(appxFile, appxPath, true);
                    }
                }

                foreach (AppxPackage package in canonicalCompdb.AppX.AppXPackages.Package)
                {
                    if (package.LicenseData != null)
                    {
                        string appxFolder = Path.Combine(appxRoot, Path.GetDirectoryName(package.Payload.PayloadItem.FirstOrDefault().Path));
                        if (!Directory.Exists(appxFolder))
                        {
                            Logging.Log($"Creating {appxFolder}");
                            _ = Directory.CreateDirectory(appxFolder);
                        }

                        string appxPath = Path.Combine(appxRoot, package.Payload.PayloadItem.FirstOrDefault().Path);
                        string appxLicensePath = Path.Combine(appxFolder, $"{appxLicenseFileMap[Path.GetFileName(appxPath)]}");
                        Logging.Log($"Writing license to {appxLicensePath}");
                        File.WriteAllText(appxLicensePath, package.LicenseData);
                    }
                }
            }

            Logging.Log($"Appx fixup applied.");
        }

        // CoreOS and Predev handling
        private static async Task CheckAndDownloadCoreOSPreDevUpdates(
                    OSSkuId ReportingSku,
                    string ReportingVersion,
                    MachineType MachineType,
                    string FlightRing,
                    string FlightingBranchName,
                    string BranchReadinessLevel,
                    string CurrentBranch,
                    string ReleaseType,
                    bool SyncCurrentVersionOnly,
                    string ContentType,
                    string Mail,
                    string Password,
                    string OutputFolder,
                    string Language,
                    string Edition)
        {
            Logging.Log("Checking for CoreOS and Predev updates...");

            CTAC ctac = new(ReportingSku, ReportingVersion, MachineType, FlightRing, FlightingBranchName, BranchReadinessLevel, CurrentBranch, ReleaseType, SyncCurrentVersionOnly, ContentType: ContentType);
            string token = string.Empty;
            if (!string.IsNullOrEmpty(Mail) && !string.IsNullOrEmpty(Password))
            {
                token = await MBIHelper.GenerateMicrosoftAccountTokenAsync(Mail, Password);
            }

            IEnumerable<UpdateData> data = await FE3Handler.GetUpdates(null, ctac, token, FileExchangeV3UpdateFilter.ProductRelease);

            if (!data.Any())
            {
                Logging.Log("No CoreOS/Predev updates found that matched the specified criteria.", Logging.LoggingLevel.Error);
            }
            else
            {
                Logging.Log($"Found {data.Count()} CoreOS/Predev update(s):");

                foreach (UpdateData update in data)
                {
                    Logging.Log("Title: " + update.Xml.LocalizedProperties.Title);
                    Logging.Log("Description: " + update.Xml.LocalizedProperties.Description);

                    await ProcessUpdateAsync(update, OutputFolder, MachineType, Language, Edition, true);
                }
            }
            Logging.Log("CoreOS/Predev updates completed.");
        }

        private static async Task CheckAndDownloadUpdates(
            OSSkuId ReportingSku,
            string ReportingVersion,
            MachineType MachineType,
            string FlightRing,
            string FlightingBranchName,
            string BranchReadinessLevel,
            string CurrentBranch,
            string ReleaseType,
            bool SyncCurrentVersionOnly,
            string ContentType,
            string Mail,
            string Password,
            string OutputFolder,
            string Language,
            string Edition)
        {
            try
            {
                // Existing method to handle normal updates...
            }
            catch (Exception ex)
            {
                Logging.Log($"Error during update check: {ex.Message}", Logging.LoggingLevel.Error);
            }
        }

        // Clean-up routine to remove temporary files
        private static void CleanupTemporaryFiles(string tempDirectory)
        {
            try
            {
                if (Directory.Exists(tempDirectory))
                {
                    Directory.Delete(tempDirectory, true);
                    Logging.Log($"Temporary files in {tempDirectory} have been deleted.");
                }
            }
            catch (Exception ex)
            {
                Logging.Log($"Error cleaning temporary files: {ex.Message}", Logging.LoggingLevel.Error);
            }
        }
    }
}
