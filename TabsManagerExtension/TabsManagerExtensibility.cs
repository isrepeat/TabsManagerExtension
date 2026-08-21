using Microsoft.VisualStudio.Extensibility;
using System.Threading;
using System.Threading.Tasks;

namespace TabsManagerExtension {
    /// <summary>
    /// Точка входа новой модели расширяемости. Расширение остаётся внутри процесса Visual Studio,
    /// чтобы существующий VSSDK-код мог продолжать работать вместе с новым Settings API.
    /// </summary>
    [VisualStudioContribution]
    public sealed class TabsManagerExtensibility : Extension {
        public override ExtensionConfiguration ExtensionConfiguration => new ExtensionConfiguration {
            RequiresInProcessHosting = true,
            LoadedWhen = ActivationConstraint.Or(
                ActivationConstraint.SolutionState(SolutionState.NoSolution),
                ActivationConstraint.SolutionState(SolutionState.Exists)
            )
        };

        protected override async Task OnInitializedAsync(VisualStudioExtensibility extensibility, CancellationToken cancellationToken) {
            await base.OnInitializedAsync(extensibility, cancellationToken);
            Helpers.Diagnostic.Logger.LogDebug("[Extensibility] Инициализация настроек Tabs Manager.");
            await Configuration.TabsManagerConfigurationService.InitializeAsync(extensibility, cancellationToken);
        }
    }
}
