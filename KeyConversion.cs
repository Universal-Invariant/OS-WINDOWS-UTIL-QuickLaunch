using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

    public class sKey
    {


        

        public static bool IsAssignableKey(Keys key)
        {
            Keys[] keys = { Keys.Escape, Keys.CapsLock, Keys.NumLock, Keys.Enter, Keys.PrintScreen, Keys.ControlKey, Keys.Alt, Keys.Menu, Keys.ShiftKey, Keys.LWin, Keys.RWin };
            if (keys.Contains(key)) return false;            
            return true;
        }


        public static string GetKeyDisplay(Keys key, bool simple = false)
        {
            var modifiers = key & Keys.Modifiers;
            key = key & Keys.KeyCode;

            if (key == null || key == Keys.None || key.ToString() == "") return "";
            var k = key.ToString().ToLower();
            if (k.Length == 2 && k[0] == 'd' && int.TryParse(k[1].ToString(), out int f)) { if (f >= 0 && f <= 9) k = f.ToString(); }
            if (modifiers.HasFlag(Keys.Shift)) k = k.ToUpper();
            var m = modifiers.ToString().Replace(" ", "").Replace(",", " + ") + " + ";
            if (simple) { m = "<" + modifiers.ToString().Replace("Shift", "S").Replace("Control", "C").Replace("Alt", "A").Replace(" ", "").Replace(",", "") + ">"; }
            return $"{((modifiers == Keys.None) ? "" : m)}{((!IsAssignableKey(key)) ? "" : (" " + k))} {((simple)?"":"(" + KeyCodeToUnicode(key) + ")")}";        
        }



    public static string KeyCodeToUnicode(Keys key)
        {
            byte[] keyboardState = new byte[255];
            bool keyboardStateStatus = GetKeyboardState(keyboardState);

            if (!keyboardStateStatus)
            {
                return "";
            }

            uint virtualKeyCode = (uint)key;
            uint scanCode = MapVirtualKey(virtualKeyCode, 0);
            IntPtr inputLocaleIdentifier = GetKeyboardLayout(0);

            StringBuilder result = new StringBuilder();
            ToUnicodeEx(virtualKeyCode, scanCode, keyboardState, result, (int)5, (uint)0, inputLocaleIdentifier);

            return result.ToString();
        }

        [DllImport("user32.dll")]
        static extern bool GetKeyboardState(byte[] lpKeyState);

        [DllImport("user32.dll")]
        static extern uint MapVirtualKey(uint uCode, uint uMapType);

        [DllImport("user32.dll")]
        static extern IntPtr GetKeyboardLayout(uint idThread);

        [DllImport("user32.dll")]
        static extern int ToUnicodeEx(uint wVirtKey, uint wScanCode, byte[] lpKeyState, [Out, MarshalAs(UnmanagedType.LPWStr)] StringBuilder pwszBuff, int cchBuff, uint wFlags, IntPtr dwhkl);
    }

