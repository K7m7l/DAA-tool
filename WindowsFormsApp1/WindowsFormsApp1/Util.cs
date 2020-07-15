using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Forms;
using Condition = System.Windows.Automation.Condition;
using HWND = System.IntPtr;

namespace WindowsFormsApp1
{
    class Util
    {
        public void Open(string values)
        {
            try
            {
                List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
                Process process = new Process();
                process.StartInfo.FileName = getValue(keyValuePairs, "ProcessName");
                process.Start();
                int pid = process.Id;
                Program.invokedProcess.Add(process);
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                LogException(ex);
            }
        }
        public void Close(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);

            try
            {
                System.Diagnostics.Process thisProcess = System.Diagnostics.Process.GetProcessesByName(getValue(keyValuePairs, "ProcessName"))[0];
                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcesses();
            
                foreach (System.Diagnostics.Process process in processes)
                {
                    if (process == thisProcess) continue;
                    System.IntPtr handle = process.MainWindowHandle;
                    if (handle == System.IntPtr.Zero) continue;
                    process.Close();
                } 
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                LogException(ex);
            }
        }
        public void Switch(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
            try
            {
                string processName = getValue(keyValuePairs, "ProcessName");
                string windowName = getValue(keyValuePairs, "WindowName");

                foreach (KeyValuePair<IntPtr, string> window in OpenWindowGetter.GetOpenWindows())
                {
                    IntPtr handler = window.Key;
                    string title = window.Value;
                    if (title.Trim().ToLower() == windowName.ToLower().Trim())
                    {
                        SetForegroundWindow(handler);
                        //SendMessage(handler, WM_SYSCOMMAND, SC_RESTORE, 0);
                        //AutomationElement windowElement = AutomationElement.FromHandle(handler);
                    }
                }
            }
            catch(Exception ex)
            {
                string e = ex.Message;
                LogException(ex);
            }

        }
        public void CheckControlExists(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
        }
        public void Click(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
            string processName = getValue(keyValuePairs, "ProcessName");
            string windowName = getValue(keyValuePairs, "WindowName");
            string automationID = getValue(keyValuePairs, "AutomationID");
            string localizedControlTypeProperty = getValue(keyValuePairs, "LocalizedControlTypeProperty");
            string controlTypeProperty = getValue(keyValuePairs, "ControlTypeProperty");
            string classNameProperty = getValue(keyValuePairs, "ClassName");
            string nameProperty = getValue(keyValuePairs, "Name");
            string X = getValue(keyValuePairs, "X");
            string Y = getValue(keyValuePairs, "Y");
            string isDouble = getValue(keyValuePairs, "isDouble");
            string isRightClick = getValue(keyValuePairs, "isRightClick");
            string isCheckBox = getValue(keyValuePairs, "isCheckBox");

            try
            {
                AutomationElement _automationRootElement, _actualAutomationElement;
                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcesses();

                HWND handler = HWND.Zero;
                foreach (KeyValuePair<IntPtr, string> window in OpenWindowGetter.GetOpenWindows())
                {
                    if (window.Value.Trim().ToLower() == windowName.ToLower().Trim())
                    {
                        handler = window.Key;
                        string title = window.Value;
                    }
                }

                if (handler != IntPtr.Zero)
                {
                    System.IntPtr handle = handler;
                    SetForegroundWindow(handle);
                }
                else
                {
                    var thisProcess = System.Diagnostics.Process.GetProcessesByName(processName)[0];
                    System.IntPtr handle = thisProcess.MainWindowHandle;
                    SetForegroundWindow(handle);
                }

                int ct = 0;

                do
                {
                    //Finds if automation element is available 
                    _automationRootElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, windowName, PropertyConditionFlags.IgnoreCase)); ++ct;
                    Thread.Sleep(100);
                }
                while (_automationRootElement == null && ct < 50);

                if (_automationRootElement == null)
                {
                    //foreach (AutomationElement child in AutomationElement.RootElement.FindAll(TreeScope.Subtree, Condition.TrueCondition))
                    //{
                    //    if (child.Current.Name.ToLower().Trim().Contains(windowName.ToLower().Trim()))
                    //    {
                    //        _automationRootElement = child; break;
                    //    }
                    //}
                    _automationRootElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, windowName, PropertyConditionFlags.IgnoreCase)); ++ct;
                }
                ct = 0;
                //Finds the control through automation ID 
                do
                {
                    if (!string.IsNullOrEmpty(automationID))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.AutomationIdProperty, automationID)); ++ct;
                    }
                    else if (!string.IsNullOrEmpty(controlTypeProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)); ++ct;
                    }
                    else if (!string.IsNullOrEmpty(localizedControlTypeProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, localizedControlTypeProperty)); ++ct;
                    }                    
                    else if (!string.IsNullOrEmpty(nameProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, nameProperty)); ++ct;
                    }
                    else
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.ClassNameProperty, classNameProperty)); ++ct;
                    }
                }
                while (_actualAutomationElement == null && ct < 50);
                //Presses the control GetInvokePattern().
                if (string.IsNullOrEmpty(isCheckBox))
                {
                    if (string.IsNullOrEmpty(X) && string.IsNullOrEmpty(Y))
                    {
                        GetInvokePattern(_actualAutomationElement).Invoke(); Thread.Sleep(100);
                    }
                    else
                    {
                        //System.Windows.Point clickablePoint = _actualAutomationElement.GetClickablePoint();
                        System.Windows.Point clickablePoint = new System.Windows.Point(Convert.ToDouble(X), Convert.ToDouble(Y));
                        System.Windows.Forms.Cursor.Position = new System.Drawing.Point((int)clickablePoint.X, (int)clickablePoint.Y);
                        if (string.IsNullOrEmpty(isRightClick))
                        {
                            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, new IntPtr());
                            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, new IntPtr());
                            if (isDouble.ToLower().Trim() == "true")
                            {
                                mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, new IntPtr());
                                mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, new IntPtr());
                            }
                        }
                        else
                        {
                            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, new IntPtr());
                            mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, new IntPtr());
                            if (isDouble.ToLower().Trim() == "true")
                            {
                                mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, new IntPtr());
                                mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, new IntPtr());
                            }
                        }
                    }
                }
                else
                {
                    var cb = _automationRootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, _actualAutomationElement));
                    ((TogglePattern)cb.GetCurrentPattern(TogglePattern.Pattern)).Toggle();
                }
            }
            catch(Exception ex)
            {
                string e = ex.Message;
                LogException(ex);
            }
        }
        public void SetText(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
            try
            {
                string processName = getValue(keyValuePairs, "ProcessName");
                string windowName = getValue(keyValuePairs, "WindowName");
                string automationID = getValue(keyValuePairs, "AutomationID");
                string localizedControlTypeProperty = getValue(keyValuePairs, "LocalizedControlTypeProperty");
                string controlTypeProperty = getValue(keyValuePairs, "ControlTypeProperty");
                string classNameProperty = getValue(keyValuePairs, "ClassName");
                string nameProperty = getValue(keyValuePairs, "Name");
                string X = getValue(keyValuePairs, "X");
                string Y = getValue(keyValuePairs, "Y");
                string text = getValue(keyValuePairs, "Text");
                string isDropDown = getValue(keyValuePairs, "isDropDown");
                string dropdownValue = getValue(keyValuePairs, "dropdownValue");
                string isListBox = getValue(keyValuePairs, "isListBox");
                string listItemIndex = getValue(keyValuePairs, "listItemIndex");

                AutomationElement _automationRootElement, _actualAutomationElement;
                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcesses();
                IntPtr handler = IntPtr.Zero;
                foreach (KeyValuePair<IntPtr, string> window in OpenWindowGetter.GetOpenWindows())
                {
                    if (window.Value.Trim().ToLower() == windowName.ToLower().Trim())
                    {
                        handler = window.Key;
                        string title = window.Value;
                    }
                }

                if (handler != IntPtr.Zero)
                {
                    System.IntPtr handle = handler;
                    SetForegroundWindow(handle);
                }
                else
                {
                    var thisProcess = System.Diagnostics.Process.GetProcessesByName(processName)[0];
                    System.IntPtr handle = thisProcess.MainWindowHandle;
                    SetForegroundWindow(handle);
                }

                int ct = 0;

                do
                {
                    //Finds if automation element is available 
                    _automationRootElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, windowName, PropertyConditionFlags.IgnoreCase)); ++ct;
                    Thread.Sleep(100);
                }
                while (_automationRootElement == null && ct < 50);

                if (_automationRootElement == null)
                {
                    //foreach (AutomationElement child in AutomationElement.RootElement.FindAll(TreeScope.Subtree, Condition.TrueCondition))
                    //{
                    //    if (child.Current.Name.ToLower().Trim().Contains(windowName.ToLower().Trim()))
                    //    {
                    //        _automationRootElement = child; break;
                    //    }
                    //}
                    _automationRootElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, windowName, PropertyConditionFlags.IgnoreCase)); ++ct;
                }
                ct = 0;
                //Finds the control through automation ID 
                do
                {
                    if (!string.IsNullOrEmpty(automationID))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.AutomationIdProperty, automationID)); ++ct;
                    }
                    else if (!string.IsNullOrEmpty(controlTypeProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)); ++ct;
                    }
                    else if (!string.IsNullOrEmpty(localizedControlTypeProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, localizedControlTypeProperty)); ++ct;
                    }
                    else if (!string.IsNullOrEmpty(nameProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, nameProperty)); ++ct;
                    }
                    else
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.ClassNameProperty, classNameProperty)); ++ct;
                    }
                }
                while (_actualAutomationElement == null && ct < 50);
                //Presses the control GetInvokePattern().
                if (string.IsNullOrEmpty(isListBox))
                {
                    if (string.IsNullOrEmpty(isDropDown))
                    {
                        if (string.IsNullOrEmpty(X) && string.IsNullOrEmpty(Y))
                        {
                            _actualAutomationElement.SetFocus();
                            ValuePattern etb = _actualAutomationElement.GetCurrentPattern(ValuePattern.Pattern) as ValuePattern;
                            etb.SetValue(text);
                        }
                        else
                        {
                            //System.Windows.Point clickablePoint = _actualAutomationElement.GetClickablePoint();
                            System.Windows.Point clickablePoint = new System.Windows.Point(Convert.ToDouble(X), Convert.ToDouble(Y));
                            System.Windows.Forms.Cursor.Position = new System.Drawing.Point((int)clickablePoint.X, (int)clickablePoint.Y);

                            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, new IntPtr());
                            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, new IntPtr());

                            SendKeys.SendWait(text);
                        }
                    }
                    else
                    {
                        SetSelectedComboBoxItem(_actualAutomationElement, dropdownValue);
                    }
                }
                else
                {
                    Object selectPattern = null;int i = 1;
                    if (string.IsNullOrEmpty(listItemIndex))
                    {
                        i = 1;
                    }
                    else
                    {
                        i = int.Parse(listItemIndex);
                    }

                    if (FindChildAt(_actualAutomationElement, 0).TryGetCurrentPattern(SelectionItemPattern.Pattern, out selectPattern))
                    {
                        (selectPattern as SelectionItemPattern).AddToSelection();
                        (selectPattern as SelectionItemPattern).Select();
                    }
                }
            }
            catch(Exception ex)
            {
                string e = ex.Message;
                LogException(ex);
            }
        }
        public void GetText(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
            string processName = getValue(keyValuePairs, "ProcessName");
            string windowName = getValue(keyValuePairs, "WindowName");
            string automationID = getValue(keyValuePairs, "AutomationID");
            string localizedControlTypeProperty = getValue(keyValuePairs, "LocalizedControlTypeProperty");
            string controlTypeProperty = getValue(keyValuePairs, "ControlTypeProperty");
            string classNameProperty = getValue(keyValuePairs, "ClassName");
            string nameProperty = getValue(keyValuePairs, "Name");
            string X = getValue(keyValuePairs, "X");
            string Y = getValue(keyValuePairs, "Y");

            try
            {
                AutomationElement _automationRootElement, _actualAutomationElement;
                System.Diagnostics.Process[] processes = System.Diagnostics.Process.GetProcesses();
                IntPtr handler = IntPtr.Zero;
                foreach (KeyValuePair<IntPtr, string> window in OpenWindowGetter.GetOpenWindows())
                {
                    if (window.Value.Trim().ToLower() == windowName.ToLower().Trim())
                    {
                        handler = window.Key;
                        string title = window.Value;
                    }
                }

                if (handler != IntPtr.Zero)
                {
                    System.IntPtr handle = handler;
                    SetForegroundWindow(handle);
                }
                else
                {
                    var thisProcess = System.Diagnostics.Process.GetProcessesByName(processName)[0];
                    System.IntPtr handle = thisProcess.MainWindowHandle;
                    SetForegroundWindow(handle);
                }

                int ct = 0;

                do
                {
                    //Finds if automation element is available 
                    _automationRootElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.NameProperty, windowName, PropertyConditionFlags.IgnoreCase)); ++ct;
                    Thread.Sleep(100);
                }
                while (_automationRootElement == null && ct < 50);

                if (_automationRootElement == null)
                {
                    //foreach (AutomationElement child in AutomationElement.RootElement.FindAll(TreeScope.Subtree, Condition.TrueCondition))
                    //{
                    //    if (child.Current.Name.ToLower().Trim().Contains(windowName.ToLower().Trim()))
                    //    {
                    //        _automationRootElement = child; break;
                    //    }
                    //}
                    _automationRootElement = AutomationElement.RootElement.FindFirst(TreeScope.Children, new PropertyCondition(AutomationElement.AutomationIdProperty, windowName, PropertyConditionFlags.IgnoreCase)); ++ct;
                }
                ct = 0;
                //Finds the control through automation ID 
                do
                {
                    if (!string.IsNullOrEmpty(automationID))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.AutomationIdProperty, automationID)); ++ct;
                    }
                    else if (!string.IsNullOrEmpty(controlTypeProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.ControlTypeProperty, ControlType.Pane)); ++ct;
                    }
                    else if (!string.IsNullOrEmpty(localizedControlTypeProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.LocalizedControlTypeProperty, localizedControlTypeProperty)); ++ct;
                    }
                    else if (!string.IsNullOrEmpty(nameProperty))
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Descendants, new PropertyCondition(AutomationElement.NameProperty, nameProperty)); ++ct;
                    }
                    else
                    {
                        _actualAutomationElement = _automationRootElement.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.ClassNameProperty, classNameProperty)); ++ct;
                    }
                }
                while (_actualAutomationElement == null && ct < 50);
                //Presses the control GetInvokePattern().
                
                if (string.IsNullOrEmpty(X) && string.IsNullOrEmpty(Y))
                {
                    string fdText = GetText(_actualAutomationElement);
                }
                else
                {
                    //System.Windows.Point clickablePoint = _actualAutomationElement.GetClickablePoint();
                    System.Windows.Point clickablePoint = new System.Windows.Point(Convert.ToDouble(X), Convert.ToDouble(Y));
                    System.Windows.Forms.Cursor.Position = new System.Drawing.Point((int)clickablePoint.X, (int)clickablePoint.Y);
                   
                    mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, new IntPtr());
                    mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, new IntPtr());
                        
                }
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                LogException(ex);
            }
        }
        public void SetTextToClipBoard(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
            try
            {
                string text = getValue(keyValuePairs, "Text");
                Clipboard.SetText(text, TextDataFormat.Rtf);
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                LogException(ex);
            }
        }
        public void GetTextFromClipBoard(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
            try
            {
                string text = getValue(keyValuePairs, "Text");
                string fdText = Clipboard.GetText(TextDataFormat.Rtf);
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                LogException(ex);
            }
        }
        public void Scroll(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
        } 
        public void Sleep(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
            try
            {
                Thread.Sleep(int.Parse(getValue(keyValuePairs, "Time")));
            }
            catch(Exception ex)
            {
                string e = ex.Message;
                Thread.Sleep(10000);
                LogException(ex);
            }
        }
        public void FindControl(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = getValues(values);
        }
        public List<KeyValuePair<string, string>> getValues(string values)
        {
            List<KeyValuePair<string, string>> keyValuePairs = new List<KeyValuePair<string, string>>();
            if (values.IndexOf(";") > -1)
            {
                List<string> vs = new List<string>();
                vs = values.Split(';').ToList();
                vs.RemoveAll(v => v.Trim() == string.Empty);
                vs.RemoveAll(v => v.Trim().IndexOf("=") < -1);

                foreach (string v in vs)
                {
                    KeyValuePair<string, string> kbp = new KeyValuePair<string, string>(v.Split('=')[0].Trim(), v.Split('=')[1].Trim());
                    keyValuePairs.Add(kbp);
                }
            }
            else
            {
                if (values.IndexOf("=") > -1)
                {
                    KeyValuePair<string, string> kbp = new KeyValuePair<string, string>(values.Split('=')[0].Trim(), values.Split('=')[1].Trim());
                    keyValuePairs.Add(kbp);
                }
            }
            return keyValuePairs;
        }
        public string getValue(List<KeyValuePair<string, string>> collection, string key)
        {
            try
            {
                return collection.Find(c => c.Key.ToLower().Trim() == key.ToLower().Trim()).Value;
            }
            catch (Exception ex)
            {
                string e = ex.Message;
                return string.Empty;
            }
        }


        [DllImport("User32.dll")]
        static extern int SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        internal static extern bool SendMessage(IntPtr hWnd, Int32 msg, Int32 wParam, Int32 lParam);
        static Int32 WM_SYSCOMMAND = 0x0112;
        static Int32 SC_RESTORE = 0xF120;

        [DllImport("user32.dll")]
        static extern bool FindWindow(IntPtr hWnd);
        public InvokePattern GetInvokePattern(AutomationElement element)
        {
            return element.GetCurrentPattern(InvokePattern.Pattern) as InvokePattern;
        }


        /// <summary>Contains functionality to get all the open windows.</summary>
        public static class OpenWindowGetter
        {
            /// <summary>Returns a dictionary that contains the handle and title of all the open windows.</summary>
            /// <returns>A dictionary that contains the handle and title of all the open windows.</returns>
            public static IDictionary<HWND, string> GetOpenWindows()
            {
                HWND shellWindow = GetShellWindow();
                Dictionary<HWND, string> windows = new Dictionary<HWND, string>();

                EnumWindows(delegate (HWND hWnd, int lParam)
                {
                    if (hWnd == shellWindow) return true;
                    if (!IsWindowVisible(hWnd)) return true;

                    int length = GetWindowTextLength(hWnd);
                    if (length == 0) return true;

                    StringBuilder builder = new StringBuilder(length);
                    GetWindowText(hWnd, builder, length + 1);

                    windows[hWnd] = builder.ToString();
                    return true;

                }, 0);

                return windows;
            }

            private delegate bool EnumWindowsProc(HWND hWnd, int lParam);

            [DllImport("USER32.DLL")]
            private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

            [DllImport("USER32.DLL")]
            private static extern int GetWindowText(HWND hWnd, StringBuilder lpString, int nMaxCount);

            [DllImport("USER32.DLL")]
            private static extern int GetWindowTextLength(HWND hWnd);

            [DllImport("USER32.DLL")]
            private static extern bool IsWindowVisible(HWND hWnd);

            [DllImport("USER32.DLL")]
            private static extern IntPtr GetShellWindow();
        }
        [DllImport("user32.dll")]
        private static extern void mouse_event(UInt32 dwFlags, UInt32 dx, UInt32 dy, UInt32 dwData, IntPtr dwExtraInfo);

        private const UInt32 MOUSEEVENTF_LEFTDOWN = 0x0002;
        private const UInt32 MOUSEEVENTF_LEFTUP = 0x0004;
        private const UInt32 MOUSEEVENTF_RIGHTDOWN = 0x0008;
        private const UInt32 MOUSEEVENTF_RIGHTUP = 0x0010;
        protected AutomationElement GetItem(AutomationElement element, string item)
        {
            AutomationElement elementList;
            CacheRequest cacheRequest = new CacheRequest();
            cacheRequest.Add(AutomationElement.NameProperty);
            cacheRequest.TreeScope = TreeScope.Element | TreeScope.Children;

            elementList = element.GetUpdatedCache(cacheRequest);

            foreach (AutomationElement child in elementList.CachedChildren)
                if (child.Cached.Name == item)
                    return child;

            return null;
        }
        public string GetText(AutomationElement element)
        {
            object patternObj;
            if (element.TryGetCurrentPattern(ValuePattern.Pattern, out patternObj))
            {
                var valuePattern = (ValuePattern)patternObj;
                return valuePattern.Current.Value;
            }
            else if (element.TryGetCurrentPattern(TextPattern.Pattern, out patternObj))
            {
                var textPattern = (TextPattern)patternObj;
                return textPattern.DocumentRange.GetText(-1).TrimEnd('\r'); // often there is an extra '\r' hanging off the end.
            }
            else
            {
                return element.Current.Name;
            }
        }

        public void logThis(string ex, string h)
        {
            string logPath = @"C:\DAA-Files\";
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            logPath = Path.Combine(logPath, "dayLog" + DateTime.Now.ToString("yyyyMMdd") + ".txt");


            try
            {
                // Check if file already exists. If yes, delete i
                if (File.Exists(logPath))
                {
                    using (var tw = new StreamWriter(logPath, true))
                    {
                        tw.WriteLine("=====================================================================================================================");
                        tw.WriteLine(ex);
                        tw.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        tw.Close();
                    }
                }
                else
                {
                    // Create a new file 
                    using (StreamWriter sw = File.CreateText(logPath))
                    {
                        sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                        sw.WriteLine(Environment.MachineName);
                        sw.WriteLine("****//\\\\*****");
                        sw.WriteLine(ex);
                        sw.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        sw.WriteLine("=====================================================================================================================");
                        sw.Close();
                    }
                }
            }
            catch (Exception Ex)
            {
                string ec = Ex.Message;
            }
        }
        public void LogException(Exception ex, [CallerMemberName] string memberName = "")
        {
            string logPath = @"C:\DAA-Files\";
            if (!Directory.Exists(logPath))
            {
                Directory.CreateDirectory(logPath);
            }

            logPath = Path.Combine(logPath, "dayLog" + DateTime.Now.ToString("yyyyMMdd") + ".txt");

            try
            {
                // Check if file already exists. If yes, delete i
                if (File.Exists(logPath))
                {
                    using (var tw = new StreamWriter(logPath, true))
                    {
                        tw.WriteLine("=====================================================================================================================");
                        tw.WriteLine("Message :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace + "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        tw.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        tw.Close();
                    }
                }
                else
                {
                    // Create a new file 
                    using (StreamWriter sw = File.CreateText(logPath))
                    {
                        sw.WriteLine("New file created: {0}", DateTime.Now.ToString());
                        sw.WriteLine(Environment.MachineName);
                        sw.WriteLine("****//\\\\*****");
                        sw.WriteLine("Message :" + ex.Message + "<br/>" + Environment.NewLine + "StackTrace :" + ex.StackTrace + "" + Environment.NewLine + "Date :" + DateTime.Now.ToString());
                        sw.WriteLine(Environment.NewLine + "-----------------------------------------------------------------------------" + Environment.NewLine);
                        sw.WriteLine("=====================================================================================================================");
                        sw.Close();
                    }
                }
            }
            catch (Exception Ex)
            {
                string ec = Ex.Message;
            }
        }
        private static WINDOWPLACEMENT GetPlacement(IntPtr hwnd)
        {
            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(placement);
            GetWindowPlacement(hwnd, ref placement);
            return placement;
        }

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowPlacement(
            IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);

        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        internal struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public ShowWindowCommands showCmd;
            public System.Drawing.Point ptMinPosition;
            public System.Drawing.Point ptMaxPosition;
            public System.Drawing.Rectangle rcNormalPosition;
        }

        internal enum ShowWindowCommands : int
        {
            Hide = 0,
            Normal = 1,
            Minimized = 2,
            Maximized = 3,
        }
        AutomationElement FindChildAt(AutomationElement parent, int index)
        {
            if (index < 0)
            {
                throw new ArgumentOutOfRangeException();
            }
            TreeWalker walker = TreeWalker.ControlViewWalker;
            AutomationElement child = walker.GetFirstChild(parent);
            for (int x = 1; x <= index; x++)
            {
                child = walker.GetNextSibling(child);
                if (child == null)
                {
                    throw new ArgumentOutOfRangeException();
                }
            }
            return child;
        }
        public void SetSelectedComboBoxItem(AutomationElement comboBox, string item)
        {
            AutomationPattern automationPatternFromElement = GetSpecifiedPattern(comboBox, "ExpandCollapsePatternIdentifiers.Pattern");

            ExpandCollapsePattern expandCollapsePattern = comboBox.GetCurrentPattern(automationPatternFromElement) as ExpandCollapsePattern;

            expandCollapsePattern.Expand();
            expandCollapsePattern.Collapse();

            AutomationElement listItem = comboBox.FindFirst(TreeScope.Subtree, new PropertyCondition(AutomationElement.NameProperty, item));

            automationPatternFromElement = GetSpecifiedPattern(listItem, "SelectionItemPatternIdentifiers.Pattern");

            SelectionItemPattern selectionItemPattern = listItem.GetCurrentPattern(automationPatternFromElement) as SelectionItemPattern;

            selectionItemPattern.Select();
        }

        private AutomationPattern GetSpecifiedPattern(AutomationElement element, string patternName)
        {
            AutomationPattern[] supportedPattern = element.GetSupportedPatterns();

            foreach (AutomationPattern pattern in supportedPattern)
            {
                if (pattern.ProgrammaticName == patternName)
                    return pattern;
            }

            return null;
        }
    }
}
