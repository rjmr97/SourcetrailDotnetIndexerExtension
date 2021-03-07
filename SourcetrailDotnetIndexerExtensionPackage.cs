using Microsoft.VisualStudio.Shell;
using System;
using System.Runtime.InteropServices;
using System.Threading;
using Task = System.Threading.Tasks.Task;

namespace SourcetrailDotnetIndexerExtension
{
    [PackageRegistration(UseManagedResourcesOnly = true, AllowsBackgroundLoading = true)]
    [Guid(PackageGuidString)]
    [ProvideOptionPage(typeof(Settings), "SourceTrail DotnetIndexer", "General", 0, 0, true)]
    [ProvideMenuResource("Menus.ctmenu", 1)]
    public sealed class SourcetrailDotnetIndexerExtensionPackage : AsyncPackage
    {
        public const string PackageGuidString = "6d80d7d6-c232-4194-8af8-13a84c23d140";

        #region Package Members

        protected override async Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);
            await GenerateCommand.InitializeAsync(this);
        }

        #endregion
    }
}
