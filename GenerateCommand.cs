using EnvDTE;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Threading;
using System;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Task = System.Threading.Tasks.Task;

namespace SourcetrailDotnetIndexerExtension
{
    internal sealed class GenerateCommand
    {
        public const int CommandId = 0x0100;

        public static readonly Guid CommandSet = new("21f3e264-c111-40ae-b6ac-46f07066d1dd");

        private SourcetrailDotnetIndexerExtensionPackage Package { get; init; }


        private GenerateCommand(SourcetrailDotnetIndexerExtensionPackage package, OleMenuCommandService commandService)
        {
            Package = package;

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }


        public static GenerateCommand Instance { get; private set; }

        public static async Task InitializeAsync(SourcetrailDotnetIndexerExtensionPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);

            OleMenuCommandService commandService = (OleMenuCommandService)await package.GetServiceAsync(typeof(IMenuCommandService));
            Instance = new GenerateCommand(package, commandService);
        }


        private async void Execute(object sender, EventArgs e)
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();


                var dte = (DTE)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(DTE));


                var window = OpenOutputWindow(dte);

                if (window == null)
                {
                    ShowMessage("Could not open output window.", null);
                    return;
                }


                var buildOutputPane = CreateAndOpenOutputPane(VSConstants.OutputWindowPaneGuid.BuildOutputPane_guid, "Build");

                if (buildOutputPane == null)
                {
                    ShowMessage("Could not open build pane on output window.", null);
                    return;
                }


                var project = GetSelectedProject();

                if (project == null)
                {
                    ShowMessage("Could not find selected project.\r\nMake sure you have a project selected on the solution explorer.", null);
                    return;
                }


                bool projectBuild = BuildProject(dte, project);

                if (projectBuild == false)
                {
                    ShowMessage("Could not build selected project.", null);
                    return;
                }


                var targetPath = GetProjectTargetPath(project);

                if (string.IsNullOrEmpty(targetPath))
                {
                    ShowMessage("Could not get selected project target path.", null);
                    return;
                }


                var generalOutputPane = CreateAndOpenOutputPane(VSConstants.OutputWindowPaneGuid.GeneralPane_guid, "General");

                if (generalOutputPane == null)
                {
                    ShowMessage("Could not open general pane on output window.", null);
                    return;
                }


                Settings settings = (Settings)Package.GetDialogPage(typeof(Settings));

                if (settings == null)
                {
                    ShowMessage("Could not get extension settings.", null);
                    return;
                }


                string outputPath = !string.IsNullOrEmpty(settings.Indexer_OutputPath)
                    ? settings.Indexer_OutputPath
                    : Path.Combine(Path.GetDirectoryName(targetPath), "sourcetrail-db");


                bool runDotnetIndexer = await RunDotnetIndexerAsync(generalOutputPane, settings, targetPath, outputPath);

                if (runDotnetIndexer == false)
                {
                    ShowMessage("Error while running DotnetIndexer.", null);
                    return;
                }


                if (settings.Sourcetrail_OpenAfterGenerating)
                {
                    string sourcetrailDBPath = Path.Combine(outputPath, Path.GetFileNameWithoutExtension(targetPath) + ".srctrlprj");

                    bool runSourcetrail = await RunSourcetrailAsync(settings, sourcetrailDBPath);

                    if (runSourcetrail == false)
                    {
                        ShowMessage("Error while opening Sourcetrail.", null);
                        return;
                    }
                }
            }
            catch (Exception exc)
            {
                ShowMessage("Unexpected error occurred.", null);
            }
        }


        private void ShowMessage(string message, string title)
        {
            VsShellUtilities.ShowMessageBox(
                Package,
                message,
                title,
                OLEMSGICON.OLEMSGICON_WARNING,
                OLEMSGBUTTON.OLEMSGBUTTON_OK,
                OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
        }


        private Window OpenOutputWindow(DTE dte)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                Window window = dte.Windows.Item(EnvDTE.Constants.vsWindowKindOutput);

                window.Activate();

                return window;
            }
            catch (Exception exc)
            {
                return null;
            }
        }

        private IVsOutputWindowPane CreateAndOpenOutputPane(Guid paneGuid, string name)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                IVsOutputWindow outWindow = (IVsOutputWindow)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsOutputWindow));
                outWindow.CreatePane(paneGuid, name, 1, 0);

                IVsOutputWindowPane pane;
                outWindow.GetPane(ref paneGuid, out pane);

                pane.Activate();
                pane.Clear();

                return pane;
            }
            catch (Exception exc)
            {
                return null;
            }
        }

        private Project GetSelectedProject()
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();


                IVsMonitorSelection monitorSelection = (IVsMonitorSelection)Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SVsShellMonitorSelection));

                monitorSelection.GetCurrentSelection(out IntPtr hierarchyPointer, out uint projectItemId, out _, out _);


                IVsHierarchy selectedHierarchy = (IVsHierarchy)Marshal.GetTypedObjectForIUnknown(hierarchyPointer, typeof(IVsHierarchy));

                selectedHierarchy.GetProperty(projectItemId, (int)__VSHPROPID.VSHPROPID_ExtObject, out object selectedObject);


                return (Project)selectedObject;
            }
            catch (Exception exc)
            {
                return null;
            }
        }

        private string GetProjectTargetPath(Project project)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                using (var projectCollection = new Microsoft.Build.Evaluation.ProjectCollection())
                {
                    var evaluationProject = new Microsoft.Build.Evaluation.Project(
                        project.FullName,
                        null,
                        null,
                        projectCollection,
                        Microsoft.Build.Evaluation.ProjectLoadSettings.Default);

                    return evaluationProject.Properties
                        .Where(x => x.Name == "TargetPath")
                        .Select(x => x.EvaluatedValue)
                        .SingleOrDefault();
                }
            }
            catch (Exception exc)
            {
                return null;
            }
        }


        private bool BuildProject(DTE dte, Project project)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();

                var solutionBuild = dte.Solution.SolutionBuild;
                var solutionConfiguration = solutionBuild.ActiveConfiguration;

                solutionBuild.BuildProject(solutionConfiguration.Name, project.UniqueName, true);

                return true;
            }
            catch (Exception exc)
            {
                return false;
            }
        }

        private async Task<bool> RunDotnetIndexerAsync(IVsOutputWindowPane generalOutputPane, Settings settings, string targetPath, string outputPath)
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();


                SynchronizationContext uiContext = SynchronizationContext.Current;


                await ThreadHelper.JoinableTaskFactory.RunAsync(async delegate
                {
                    await TaskScheduler.Default;

                    using (System.Diagnostics.Process exeProcess = new System.Diagnostics.Process())
                    {
                        exeProcess.StartInfo = new ProcessStartInfo()
                        {
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            WindowStyle = ProcessWindowStyle.Hidden,
                            RedirectStandardOutput = true,
                            FileName = settings.Indexer_ExecutablePath,
                            Arguments = GetRunDotnetIndexerArguments(settings, targetPath, outputPath)
                        };

                        exeProcess.OutputDataReceived += (sender2, e2) =>
                        {
                            uiContext.Send((x) =>
                            {
                                ThreadHelper.ThrowIfNotOnUIThread();
                                generalOutputPane.OutputString(e2.Data + Environment.NewLine);
                            }, null);
                        };

                        exeProcess.Start();

                        exeProcess.BeginOutputReadLine();

                        exeProcess.WaitForExit();
                    }
                });


                return true;
            }
            catch (Exception exc)
            {
                return false;
            }
        }

        private string GetRunDotnetIndexerArguments(Settings optionPageGrid, string targetPath, string outputPath)
        {
            StringBuilder argsBuilder = new StringBuilder();

            argsBuilder.Append(" -i ").Append("\"" + targetPath + "\"");

            if (optionPageGrid.Indexer_SearchPaths != null)
            {
                foreach (var searchPath in optionPageGrid.Indexer_SearchPaths)
                {
                    argsBuilder.Append(" -s ").Append("\"" + searchPath + "\"");
                }
            }

            if (optionPageGrid.Indexer_NameFilters != null)
            {
                foreach (var nameFilter in optionPageGrid.Indexer_NameFilters)
                {
                    argsBuilder.Append(" -f ").Append("\"" + nameFilter + "\"");
                }
            }

            argsBuilder.Append(" -o ").Append("\"" + outputPath + "\"");

            if (!string.IsNullOrEmpty(optionPageGrid.Indexer_OutputFilename))
            {
                argsBuilder.Append(" -of ").Append("\"" + optionPageGrid.Indexer_OutputFilename + "\"");
            }

            return argsBuilder.ToString();
        }


        private async Task<bool> RunSourcetrailAsync(Settings settings, string sourcetrailDBPath)
        {
            try
            {
                await ThreadHelper.JoinableTaskFactory.RunAsync(async delegate
                {
                    await TaskScheduler.Default;

                    using (System.Diagnostics.Process exeProcess = new System.Diagnostics.Process())
                    {
                        exeProcess.StartInfo = new ProcessStartInfo()
                        {
                            UseShellExecute = false,
                            CreateNoWindow = true,
                            WindowStyle = ProcessWindowStyle.Hidden,
                            RedirectStandardOutput = true,
                            FileName = settings.Sourcetrail_ExecutablePath,
                            Arguments = "--project-file \"" + sourcetrailDBPath + "\""
                        };

                        exeProcess.Start();
                    }
                });


                return true;
            }
            catch (Exception exc)
            {
                return false;
            }
        }
    }
}
