using System;
using System.ComponentModel.Composition;
using System.ComponentModel.Design;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Runtime.InteropServices;
using System.Threading;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Workspace;
using Microsoft.VisualStudio.Workspace.Indexing;
using Microsoft.VisualStudio.Workspace.VSIntegration;
using Microsoft.VisualStudio.Workspace.VSIntegration.Contracts;
using Microsoft.Win32;

namespace VSIXProject
{
    /// <summary>
    /// This is the class that implements the package exposed by this assembly.
    /// </summary>
    /// <remarks>
    /// <para>
    /// The minimum requirement for a class to be considered a valid package for Visual Studio
    /// is to implement the IVsPackage interface and register itself with the shell.
    /// This package uses the helper classes defined inside the Managed Package Framework (MPF)
    /// to do it: it derives from the Package class that provides the implementation of the
    /// IVsPackage interface and uses the registration attributes defined in the framework to
    /// register itself and its components with the shell. These attributes tell the pkgdef creation
    /// utility what data to put into .pkgdef file.
    /// </para>
    /// <para>
    /// To get loaded into VS, the package must be referred by &lt;Asset Type="Microsoft.VisualStudio.VsPackage" ...&gt; in .vsixmanifest file.
    /// </para>
    /// </remarks>
    [PackageRegistration(UseManagedResourcesOnly = true)]
    [InstalledProductRegistration("#110", "#112", "1.0", IconResourceID = 400)] // Info on this package for Help/About
    [Guid(VSPackage.PackageGuidString)]
    [SuppressMessage("StyleCop.CSharp.DocumentationRules", "SA1650:ElementDocumentationMustBeSpelledCorrectly", Justification = "pkgdef, VS and vsixmanifest are valid VS terms")]
    [ProvideAutoLoad(VSConstants.UICONTEXT.FolderOpened_string)]
    public sealed class VSPackage : AsyncPackage
    {
        public IVsFolderWorkspaceService FolderWorkspaceService { get; private set; }

        /// <summary>
        /// VSPackage GUID string.
        /// </summary>
        public const string PackageGuidString = "f4924691-d8b1-4480-99ab-099a9bcd98c0";

        /// <summary>
        /// Initializes a new instance of the <see cref="VSPackage"/> class.
        /// </summary>
        public VSPackage()
        {
            // Inside this method you can place any initialization code that does not require
            // any Visual Studio service because at this point the package object is created but
            // not sited yet inside Visual Studio environment. The place to do all the other
            // initialization is the Initialize method.
        }

        protected override async System.Threading.Tasks.Task InitializeAsync(CancellationToken cancellationToken, IProgress<ServiceProgressData> progress)
        {
            await base.InitializeAsync(cancellationToken, progress);

            var componentModel = (IComponentModel)await this.GetServiceAsync(typeof(SComponentModel));
            this.FolderWorkspaceService = componentModel.GetService<IVsFolderWorkspaceService>();

            this.FolderWorkspaceService.OnActiveWorkspaceChanged += OnActiveWorkspaceChangedAsync;

            await this.WriteWorkspaceInfoAsync(FolderWorkspaceService.CurrentWorkspace);

            var indexService = this.FolderWorkspaceService.CurrentWorkspace.GetIndexWorkspaceService();
            indexService.OnFileScannerCompleted += OnFileScannerCompletedAsync;

            if ((indexService.State != IndexWorkspaceState.FileScanning) && (indexService.State != IndexWorkspaceState.FileSystem))
            {
                await this.OnFileScannerCompletedAsync(indexService, null);
            }
        }

        private async System.Threading.Tasks.Task OnActiveWorkspaceChangedAsync(object sender, EventArgs e)
        {
            if (sender is IVsFolderWorkspaceService folderWorkspaceService)
            {
                if (folderWorkspaceService.CurrentWorkspace == null)
                {
                    await this.WriteToOutputWindowAsync("Workspace closed." + Environment.NewLine);
                }
                else
                {
                    await this.WriteToOutputWindowAsync("Workspace opened." + Environment.NewLine);
                    await this.WriteWorkspaceInfoAsync(folderWorkspaceService.CurrentWorkspace);

                    var indexService = folderWorkspaceService.CurrentWorkspace.GetIndexWorkspaceService();
                    indexService.OnFileScannerCompleted += OnFileScannerCompletedAsync;
                }
            }
        }

        private async System.Threading.Tasks.Task OnFileScannerCompletedAsync(object sender, FileScannerEventArgs e)
        {
            if (sender is IIndexWorkspaceService indexService)
            {
                await this.WriteToOutputWindowAsync("File scanner status: " + indexService.State + Environment.NewLine);

                foreach (var file in await this.FolderWorkspaceService.CurrentWorkspace.GetFilesAsync(string.Empty, recursive: true))
                {
                    var fileReferencesResult = await indexService.GetFileReferencesAsync(file);

                    foreach (var result in fileReferencesResult)
                    {
                        await this.WriteToOutputWindowAsync("Path: " + result.Path + Environment.NewLine);
                        await this.WriteToOutputWindowAsync("Context: " + result.Context + Environment.NewLine);
                        await this.WriteToOutputWindowAsync("Target: " + result.Target + Environment.NewLine);
                        await this.WriteToOutputWindowAsync("Type: " + result.Type + Environment.NewLine);
                    }
                }
            }
        }

        private async System.Threading.Tasks.Task WriteWorkspaceInfoAsync(IWorkspace workspace)
        {
            await this.WriteToOutputWindowAsync("Workspace location: " + workspace.Location + Environment.NewLine);

            await this.WriteToOutputWindowAsync("Directories:" + Environment.NewLine);
            foreach (var directory in await workspace.GetDirectoriesAsync(string.Empty, recursive: true))
            {
                await this.WriteToOutputWindowAsync("\t" + directory + Environment.NewLine);
            }

            await this.WriteToOutputWindowAsync(Environment.NewLine);
            await this.WriteToOutputWindowAsync("Files:" + Environment.NewLine);

            foreach (var file in await workspace.GetFilesAsync(string.Empty, recursive: true))
            {
                await this.WriteToOutputWindowAsync("\t" + file + Environment.NewLine);
            }

            await this.WriteToOutputWindowAsync(Environment.NewLine);
        }

        private async System.Threading.Tasks.Task WriteToOutputWindowAsync(string output)
        {
            await this.JoinableTaskFactory.SwitchToMainThreadAsync();

            var outputWindow = await this.GetServiceAsync(typeof(SVsOutputWindow)) as IVsOutputWindow;

            Guid paneGuid = VSConstants.OutputWindowPaneGuid.GeneralPane_guid;
            int hr = outputWindow.GetPane(ref paneGuid, out IVsOutputWindowPane pane);

            if (ErrorHandler.Failed(hr) || (pane == null))
            {
                if (ErrorHandler.Succeeded(outputWindow.CreatePane(ref paneGuid, "General", fInitVisible: 1, fClearWithSolution: 1)))
                {
                    hr = outputWindow.GetPane(ref paneGuid, out pane);
                }
            }

            if (ErrorHandler.Succeeded(hr))
            {
                pane?.Activate();
                pane?.OutputString(output);
            }
        }
    }
}
