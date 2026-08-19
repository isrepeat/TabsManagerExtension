using System;
using System.Linq;
using System.Windows;
using System.Windows.Input;
using System.Windows.Threading;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio;
using Microsoft.VisualStudio.OLE.Interop;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.TextManager.Interop;


namespace TabsManagerExtension.VsShell.TextEditor {
    public class TextEditorCommandFilter : OleCommandFilterBase {

        public event Action<Guid, uint, IntPtr>? CommandIntercepted;
        public event Action<Guid, uint, IntPtr>? CommandPassedThrough;
        public bool IsEnabled { get; set; } = false;
        public bool IsSpaceCommandEnabled { get; set; }

        private readonly IReadOnlyCollection<VSConstants.VSStd2KCmdID> _trackedStd2kCommands;
        private readonly IReadOnlyCollection<VSConstants.VSStd97CmdID> _trackedStd97Commands;
        //private readonly IReadOnlyCollection<VSConstants.VSStd97CmdID> _trackedMappedStd97FromStd2kCommands;

        //private readonly Dictionary<VSConstants.VSStd2KCmdID, VSConstants.VSStd97CmdID> _2kTo97 = new();
        //private readonly Dictionary<VSConstants.VSStd97CmdID, VSConstants.VSStd2KCmdID> _97To2k = new();

        // Empty Ctor may be used for tracking passed through commands. 
        public TextEditorCommandFilter() {
            _trackedStd2kCommands = new HashSet<VSConstants.VSStd2KCmdID>();
            _trackedStd97Commands = new HashSet<VSConstants.VSStd97CmdID>();
        }

        public TextEditorCommandFilter(
            IEnumerable<VSConstants.VSStd2KCmdID> trackedStd2kCommands,
            IReadOnlyCollection<VSConstants.VSStd97CmdID> trackedStd97Commands
            ) {
            _trackedStd2kCommands = new HashSet<VSConstants.VSStd2KCmdID>(trackedStd2kCommands);
            _trackedStd97Commands = new HashSet<VSConstants.VSStd97CmdID>(trackedStd97Commands);
            //_trackedMappedStd97FromStd2kCommands = VsShell.VsCommandMapper.GetMappedStd97FromStd2kCommands(trackedStd2kCommands);
        }

        //
        // OleCommandFilterBase
        //
        protected override bool TryHandleCommand(Guid cmdGroup, uint cmdId, IntPtr inputArgument) {
            if (!this.IsEnabled) {
                return false;
            }

            if (cmdGroup == VSConstants.VSStd2K && Enum.IsDefined(typeof(VSConstants.VSStd2KCmdID), (int)cmdId)) {
                var std2kCmd = (VSConstants.VSStd2KCmdID)cmdId;
                if (std2kCmd == VSConstants.VSStd2KCmdID.TYPECHAR) {
                    return this.IsSpaceCommandEnabled && IsSpaceInput(inputArgument);
                }

                return _trackedStd2kCommands.Contains(std2kCmd);
            }

            if (cmdGroup == VSConstants.GUID_VSStandardCommandSet97 && Enum.IsDefined(typeof(VSConstants.VSStd97CmdID), (int)cmdId)) {
                var std97Cmd = (VSConstants.VSStd97CmdID)cmdId;
                //return _trackedStd97Commands.Contains(std97Cmd) || _trackedMappedStd97FromStd2kCommands.Contains(std97Cmd);
                return _trackedStd97Commands.Contains(std97Cmd);
            }

            return false;
        }


        protected override void OnCommandIntercepted(Guid cmdGroup, uint cmdId, IntPtr inputArgument) {
            this.CommandIntercepted?.Invoke(cmdGroup, cmdId, inputArgument);
        }


        protected override void OnCommandPassedThrough(Guid cmdGroup, uint cmdId, IntPtr inputArgument) {
            this.CommandPassedThrough?.Invoke(cmdGroup, cmdId, inputArgument);
        }


        //
        // Api
        //
        public Key? TryResolveKey(Guid cmdGroup, uint cmdId, IntPtr inputArgument) {
            if (cmdGroup == VSConstants.VSStd2K && Enum.IsDefined(typeof(VSConstants.VSStd2KCmdID), (int)cmdId)) {
                var std2kCmd = (VSConstants.VSStd2KCmdID)cmdId;
                if (std2kCmd == VSConstants.VSStd2KCmdID.TYPECHAR && IsSpaceInput(inputArgument)) {
                    return Key.Space;
                }

                return VsShell.VsCommandMapper.TryMapStd2kToKey(std2kCmd);
            }

            if (cmdGroup == VSConstants.GUID_VSStandardCommandSet97 && Enum.IsDefined(typeof(VSConstants.VSStd97CmdID), (int)cmdId)) {
                var std97Cmd = (VSConstants.VSStd97CmdID)cmdId;

                //if (_trackedStd97Commands.Contains(std97Cmd) ||  _trackedMappedStd97FromStd2kCommands.Contains(std97Cmd)) {
                if (_trackedStd97Commands.Contains(std97Cmd)) {
                    if (VsShell.VsCommandMapper.TryMapStd97ToStd2kCommand(std97Cmd, out var std2kCmd)) {
                        return VsShell.VsCommandMapper.TryMapStd2kToKey(std2kCmd);
                    }
                }
            }

            return null;
        }

        private static bool IsSpaceInput(IntPtr inputArgument) {
            if (inputArgument == IntPtr.Zero) {
                return false;
            }

            var value = Marshal.GetObjectForNativeVariant(inputArgument);
            return value switch {
                char character => character == ' ',
                ushort characterCode => characterCode == ' ',
                short characterCode => characterCode == ' ',
                string text => text == " ",
                _ => false
            };
        }
    }
}
