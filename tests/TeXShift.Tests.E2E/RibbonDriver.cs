using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Automation;
using System.Xml.Linq;

namespace TeXShift.Tests.E2E
{
    /// <summary>Invokes the installed add-in through OneNote's accessibility provider.</summary>
    internal sealed class RibbonDriver
    {
        private delegate bool EnumWindowCallback(IntPtr window, IntPtr parameter);
        [DllImport("user32.dll")]
        private static extern bool EnumWindows(EnumWindowCallback callback, IntPtr parameter);
        [DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
        [DllImport("user32.dll")]
        private static extern bool IsWindowVisible(IntPtr window);
        [DllImport("user32.dll")]
        private static extern IntPtr GetWindow(IntPtr window, uint command);
        private readonly AutomationElement _window;
        private readonly IntPtr _windowHandle;
        private AutomationElement _button;
        private AutomationElement _originalTab;
        public bool CommandPending { get; private set; }

        public RibbonDriver(IntPtr window) { _windowHandle = window; _window = AutomationElement.FromHandle(window); }

        public void Prepare(string command)
        {
            var allTabs = _window.FindAll(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem));
            _originalTab = allTabs.Cast<AutomationElement>().FirstOrDefault(tab =>
                tab.TryGetCurrentPattern(SelectionItemPattern.Pattern, out var pattern) &&
                ((SelectionItemPattern)pattern).Current.IsSelected);
            var tabs = _window.FindAll(TreeScope.Descendants,
                new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.TabItem),
                    new PropertyCondition(AutomationElement.NameProperty, "TeXShift")));
            if (tabs.Count != 1)
                throw new InvalidOperationException($"Expected one TeXShift Ribbon tab; found {tabs.Count}. Check add-in registration and restart OneNote.");
            ((SelectionItemPattern)tabs[0].GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
            var names = command == "paste-convert" ? new[] { "Paste and Convert", "粘贴并转换" }
                : command == "reverse" ? new[] { "Reverse Convert", "反向转换" } : new[] { "Convert", "一键转换" };
            var deadline = DateTime.UtcNow.AddSeconds(10);
            do
            {
                var content = _window.FindFirst(TreeScope.Descendants,
                    new AndCondition(new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Custom),
                        new PropertyCondition(AutomationElement.NameProperty, "TeXShift")));
                var buttons = content?.FindAll(TreeScope.Descendants,
                    new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Button));
                _button = buttons?.Cast<AutomationElement>().SingleOrDefault(b => names.Contains(b.Current.Name));
                if (_button != null && _button.Current.IsEnabled) return;
                Thread.Sleep(100);
            } while (DateTime.UtcNow < deadline);
            throw new TimeoutException("TeXShift conversion button was not available.");
        }

        public void Restore()
        {
            if (_originalTab != null)
                ((SelectionItemPattern)_originalTab.GetCurrentPattern(SelectionItemPattern.Pattern)).Select();
        }

        public void Undo()
        {
            var undo = _window.FindFirst(TreeScope.Descendants,
                new PropertyCondition(AutomationElement.AutomationIdProperty, "UndoButton"));
            if (undo == null || !undo.Current.IsEnabled)
                throw new InvalidOperationException("OneNote Undo is unavailable.");
            ((InvokePattern)undo.GetCurrentPattern(InvokePattern.Pattern)).Invoke();
        }

        public async Task InvokeAsync(string command, string pageId, string output)
        {
            var observation = new RibbonCommandObservation(command, pageId, DateTime.UtcNow);
            var elapsed = Stopwatch.StartNew();
            CommandPending = true;
            // A modal dialog can keep Invoke blocked. Observe and dismiss it from another MTA thread.
            var invoke = Task.Run(() => ((InvokePattern)_button.GetCurrentPattern(InvokePattern.Pattern)).Invoke());
            var visibleDialogs = new Dictionary<IntPtr, long>();
            string dialogText = null;
            var deadline = DateTime.UtcNow.AddSeconds(90);
            try
            {
                while (DateTime.UtcNow < deadline)
                {
                    if (invoke.IsFaulted) await invoke.ConfigureAwait(false);
                    observation.ReadLogs();
                    if (observation.ProcessId.HasValue)
                        dialogText = DismissErrorDialog(observation.ProcessId.Value, elapsed, visibleDialogs) ?? dialogText;
                    if (observation.Completed)
                    {
                        CommandPending = false;
                        if (await Task.WhenAny(invoke, Task.Delay(5000)).ConfigureAwait(false) != invoke)
                            throw new TimeoutException("The command completed but the UI Automation invocation did not return.");
                        await invoke.ConfigureAwait(false);
                        if (!observation.Success)
                            throw new InvalidOperationException("TeXShift Ribbon command failed. " +
                                (dialogText == null ? "See F09_Ribbon_Command.xml." :
                                    (string)XElement.Parse(dialogText).Elements("Control").FirstOrDefault(c =>
                                        (string)c.Attribute("type") == "ControlType.Text")?.Attribute("name")
                                    + " See F11_Error_Dialog.txt."));
                        return;
                    }
                    await Task.Delay(150).ConfigureAwait(false);
                }
                throw new TimeoutException("No completed TeXShift Ribbon command was observed. Check F09_Ribbon_Command.xml.");
            }
            finally
            {
                observation.Save(Path.Combine(output, "F09_Ribbon_Command.xml"));
                if (dialogText != null) File.WriteAllText(Path.Combine(output, "F11_Error_Dialog.txt"), dialogText);
            }
        }

        private string DismissErrorDialog(int processId, Stopwatch elapsed, Dictionary<IntPtr, long> visibleDialogs)
        {
            // Enumerating the desktop through UIA can block on OneNote while its Ribbon callback
            // waits for the modal dialog. Scope native handles to the observed add-in process first.
            var windows = new List<IntPtr>();
            EnumWindows((window, unused) =>
            {
                GetWindowThreadProcessId(window, out uint pid);
                if (pid == processId && IsWindowVisible(window)) windows.Add(window);
                return true;
            }, IntPtr.Zero);
            foreach (var window in windows)
            {
                // WPF's UIA provider can become available after the window is already visible.
                if (!visibleDialogs.ContainsKey(window)) visibleDialogs.Add(window, elapsed.ElapsedMilliseconds);
                var dialog = AutomationElement.FromHandle(window);
                if (dialog.Current.AutomationId != "TeXShift.ErrorDialog" &&
                    !(dialog.Current.ClassName == "#32770" && dialog.Current.Name.StartsWith("TeXShift", StringComparison.Ordinal))) continue;
                var description = Describe(dialog);
                description.SetAttributeValue("observedAfterMs", visibleDialogs[window]);
                description.SetAttributeValue("owner", GetWindow(window, 4).ToInt64());
                description.SetAttributeValue("expectedOwner", _windowHandle.ToInt64());
                description.SetAttributeValue("ownedByOneNote", GetWindow(window, 4) == _windowHandle);
                var text = description.ToString();
                ((WindowPattern)dialog.GetCurrentPattern(WindowPattern.Pattern)).Close();
                return text;
            }
            return null;
        }

        public XElement Snapshot() => Describe(_window);

        private static XElement Describe(AutomationElement root)
        {
            // Read document content through OneNote COM. Its UIA TextPattern.GetText can fault
            // inside UIAutomationCore; accessibility properties remain sufficient for UI diagnosis.
            var controls = root.FindAll(TreeScope.Descendants, Condition.TrueCondition);
            return new XElement("UI", new XAttribute("title", root.Current.Name),
                new XAttribute("id", root.Current.AutomationId ?? ""),
                controls.Cast<AutomationElement>().Select(control =>
                {
                    var info = control.Current;
                    return new XElement("Control", new XAttribute("type", info.ControlType.ProgrammaticName),
                        new XAttribute("id", info.AutomationId ?? ""), new XAttribute("name", info.Name ?? ""),
                        new XAttribute("handle", info.NativeWindowHandle), new XAttribute("class", info.ClassName ?? ""),
                        new XAttribute("focus", info.HasKeyboardFocus),
                        new XAttribute("patterns", string.Join(",", control.GetSupportedPatterns().Select(p => p.ProgrammaticName))),
                        control.TryGetCurrentPattern(ValuePattern.Pattern, out var value)
                            ? new XAttribute("value", ((ValuePattern)value).Current.Value ?? "") : null);
                }));
        }
    }
}
