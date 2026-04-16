using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;

namespace WinAutomator
{
    public class QaKeyboardForm : Form
    {
        private class KeyVisual
        {
            public Keys[] MatchCodes;
            public string Label;
            public float WidthMultiplier;
            public Panel Panel;
            public int Clicks;
        }

        private List<List<KeyVisual>> layoutRows;

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        public QaKeyboardForm()
        {
            this.Text = "Project Aura - Keyboard Test";
            this.Size = new Size(1100, 520);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizeBox = false;
            this.KeyPreview = true;
            this.BackColor = Color.FromArgb(28, 28, 28);
            this.RightToLeft = RightToLeft.No;

            // Paint border
            this.Paint += (s, e) => {
                using (Pen p = new Pen(Color.FromArgb(70, 70, 70), 2))
                    e.Graphics.DrawRectangle(p, 0, 0, this.Width - 1, this.Height - 1);
            };

            // Custom Title Bar
            Panel titleBar = new Panel() { Dock = DockStyle.Top, Height = 40, BackColor = Color.FromArgb(20, 20, 20), Cursor = Cursors.Default };
            titleBar.MouseDown += (s, e) => {
                if (e.Button == MouseButtons.Left) {
                    ReleaseCapture();
                    SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
                }
            };
            Label lblTitle = new Label() { Text = "Project Aura - בדיקת מקלדת", ForeColor = Color.Cyan, Font = new Font("Segoe UI", 10, FontStyle.Bold), Location = new Point(10, 10), AutoSize = true, BackColor = Color.Transparent };
            titleBar.Controls.Add(lblTitle);
            
            Label lblClose = new Label() { Text = "✕", Font = new Font("Segoe UI", 12, FontStyle.Bold), Location = new Point(this.Width - 30, 8), AutoSize = true, ForeColor = Color.DarkGray, Cursor = Cursors.Hand, BackColor = Color.Transparent };
            lblClose.MouseEnter += (s, e) => lblClose.ForeColor = Color.LightCoral;
            lblClose.MouseLeave += (s, e) => lblClose.ForeColor = Color.DarkGray;
            lblClose.Click += (s, e) => this.Close();
            titleBar.Controls.Add(lblClose);
            this.Controls.Add(titleBar);

            Label lblInstr = new Label() { 
                Text = "הקש על מקשי המקלדת. צבע המקש ישתנה בהתאם למספר הלחיצות:\n" +
                       "אפור (לא נלחץ)   ➔   ירוק (פעם 1)   ➔   סגול (פעם 2)   ➔   תכלת (פעם 3)", 
                Location = new Point(0, 40), Width = this.Width, Height = 60, ForeColor = Color.White, 
                Font = new Font("Segoe UI", 12, FontStyle.Regular), TextAlign = ContentAlignment.MiddleCenter,
                RightToLeft = RightToLeft.Yes
            };
            this.Controls.Add(lblInstr);

            Panel keyboardContainer = new Panel() { 
                Location = new Point(20, 100), Size = new Size(1060, 340),
                AutoScroll = false,
                BackColor = Color.Transparent
            };
            this.Controls.Add(keyboardContainer);

            BuildKeyboardLayout();
            RenderKeyboard(keyboardContainer);

            Label btnDone = new Label() { 
                Text = "סיימתי ובסדר ✓", Location = new Point(this.Width / 2 - 100, 460), Size = new Size(200, 40), 
                BackColor = Color.SeaGreen, ForeColor = Color.White, 
                Font = new Font("Segoe UI", 12, FontStyle.Bold), Cursor = Cursors.Hand,
                TextAlign = ContentAlignment.MiddleCenter
            };
            btnDone.Click += (s, e) => this.Close();
            this.Controls.Add(btnDone);

            this.KeyDown += QaKeyboardForm_KeyDown;
        }

        private void BuildKeyboardLayout()
        {
            layoutRows = new List<List<KeyVisual>>();

            // Row 0: Function Keys
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Esc", MatchCodes=new[]{Keys.Escape}, WidthMultiplier=1 },
                new KeyVisual { Label="", MatchCodes=new Keys[0], WidthMultiplier=0.5f }, // Gap
                new KeyVisual { Label="F1", MatchCodes=new[]{Keys.F1}, WidthMultiplier=1 },
                new KeyVisual { Label="F2", MatchCodes=new[]{Keys.F2}, WidthMultiplier=1 },
                new KeyVisual { Label="F3", MatchCodes=new[]{Keys.F3}, WidthMultiplier=1 },
                new KeyVisual { Label="F4", MatchCodes=new[]{Keys.F4}, WidthMultiplier=1 },
                new KeyVisual { Label="", MatchCodes=new Keys[0], WidthMultiplier=0.3f }, // Gap
                new KeyVisual { Label="F5", MatchCodes=new[]{Keys.F5}, WidthMultiplier=1 },
                new KeyVisual { Label="F6", MatchCodes=new[]{Keys.F6}, WidthMultiplier=1 },
                new KeyVisual { Label="F7", MatchCodes=new[]{Keys.F7}, WidthMultiplier=1 },
                new KeyVisual { Label="F8", MatchCodes=new[]{Keys.F8}, WidthMultiplier=1 },
                new KeyVisual { Label="", MatchCodes=new Keys[0], WidthMultiplier=0.3f }, // Gap
                new KeyVisual { Label="F9", MatchCodes=new[]{Keys.F9}, WidthMultiplier=1 },
                new KeyVisual { Label="F10", MatchCodes=new[]{Keys.F10}, WidthMultiplier=1 },
                new KeyVisual { Label="F11", MatchCodes=new[]{Keys.F11}, WidthMultiplier=1 },
                new KeyVisual { Label="F12", MatchCodes=new[]{Keys.F12}, WidthMultiplier=1 }
            });

            // Row 1: Numbers
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="~", MatchCodes=new[]{Keys.Oem3, Keys.Oemtilde}, WidthMultiplier=1 },
                new KeyVisual { Label="1", MatchCodes=new[]{Keys.D1}, WidthMultiplier=1 },
                new KeyVisual { Label="2", MatchCodes=new[]{Keys.D2}, WidthMultiplier=1 },
                new KeyVisual { Label="3", MatchCodes=new[]{Keys.D3}, WidthMultiplier=1 },
                new KeyVisual { Label="4", MatchCodes=new[]{Keys.D4}, WidthMultiplier=1 },
                new KeyVisual { Label="5", MatchCodes=new[]{Keys.D5}, WidthMultiplier=1 },
                new KeyVisual { Label="6", MatchCodes=new[]{Keys.D6}, WidthMultiplier=1 },
                new KeyVisual { Label="7", MatchCodes=new[]{Keys.D7}, WidthMultiplier=1 },
                new KeyVisual { Label="8", MatchCodes=new[]{Keys.D8}, WidthMultiplier=1 },
                new KeyVisual { Label="9", MatchCodes=new[]{Keys.D9}, WidthMultiplier=1 },
                new KeyVisual { Label="0", MatchCodes=new[]{Keys.D0}, WidthMultiplier=1 },
                new KeyVisual { Label="-", MatchCodes=new[]{Keys.OemMinus}, WidthMultiplier=1 },
                new KeyVisual { Label="=", MatchCodes=new[]{Keys.Oemplus}, WidthMultiplier=1 },
                new KeyVisual { Label="Back", MatchCodes=new[]{Keys.Back}, WidthMultiplier=2 }
            });

            // Row 2: QWERTY
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Tab", MatchCodes=new[]{Keys.Tab}, WidthMultiplier=1.5f },
                new KeyVisual { Label="Q", MatchCodes=new[]{Keys.Q}, WidthMultiplier=1 },
                new KeyVisual { Label="W", MatchCodes=new[]{Keys.W}, WidthMultiplier=1 },
                new KeyVisual { Label="E", MatchCodes=new[]{Keys.E}, WidthMultiplier=1 },
                new KeyVisual { Label="R", MatchCodes=new[]{Keys.R}, WidthMultiplier=1 },
                new KeyVisual { Label="T", MatchCodes=new[]{Keys.T}, WidthMultiplier=1 },
                new KeyVisual { Label="Y", MatchCodes=new[]{Keys.Y}, WidthMultiplier=1 },
                new KeyVisual { Label="U", MatchCodes=new[]{Keys.U}, WidthMultiplier=1 },
                new KeyVisual { Label="I", MatchCodes=new[]{Keys.I}, WidthMultiplier=1 },
                new KeyVisual { Label="O", MatchCodes=new[]{Keys.O}, WidthMultiplier=1 },
                new KeyVisual { Label="P", MatchCodes=new[]{Keys.P}, WidthMultiplier=1 },
                new KeyVisual { Label="[", MatchCodes=new[]{Keys.OemOpenBrackets}, WidthMultiplier=1 },
                new KeyVisual { Label="]", MatchCodes=new[]{Keys.Oem6, Keys.OemCloseBrackets}, WidthMultiplier=1 },
                new KeyVisual { Label="\\", MatchCodes=new[]{Keys.Oem5, Keys.OemPipe}, WidthMultiplier=1.5f }
            });

            // Row 3: ASDF
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Caps", MatchCodes=new[]{Keys.Capital, Keys.CapsLock}, WidthMultiplier=1.8f },
                new KeyVisual { Label="A", MatchCodes=new[]{Keys.A}, WidthMultiplier=1 },
                new KeyVisual { Label="S", MatchCodes=new[]{Keys.S}, WidthMultiplier=1 },
                new KeyVisual { Label="D", MatchCodes=new[]{Keys.D}, WidthMultiplier=1 },
                new KeyVisual { Label="F", MatchCodes=new[]{Keys.F}, WidthMultiplier=1 },
                new KeyVisual { Label="G", MatchCodes=new[]{Keys.G}, WidthMultiplier=1 },
                new KeyVisual { Label="H", MatchCodes=new[]{Keys.H}, WidthMultiplier=1 },
                new KeyVisual { Label="J", MatchCodes=new[]{Keys.J}, WidthMultiplier=1 },
                new KeyVisual { Label="K", MatchCodes=new[]{Keys.K}, WidthMultiplier=1 },
                new KeyVisual { Label="L", MatchCodes=new[]{Keys.L}, WidthMultiplier=1 },
                new KeyVisual { Label=";", MatchCodes=new[]{Keys.Oem1, Keys.OemSemicolon}, WidthMultiplier=1 },
                new KeyVisual { Label="'", MatchCodes=new[]{Keys.Oem7, Keys.OemQuotes}, WidthMultiplier=1 },
                new KeyVisual { Label="Enter", MatchCodes=new[]{Keys.Enter}, WidthMultiplier=2.2f }
            });

            // Row 4: ZXCV  – use only L/R specific shift codes
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Shift", MatchCodes=new[]{Keys.LShiftKey}, WidthMultiplier=2.2f },
                new KeyVisual { Label="Z", MatchCodes=new[]{Keys.Z}, WidthMultiplier=1 },
                new KeyVisual { Label="X", MatchCodes=new[]{Keys.X}, WidthMultiplier=1 },
                new KeyVisual { Label="C", MatchCodes=new[]{Keys.C}, WidthMultiplier=1 },
                new KeyVisual { Label="V", MatchCodes=new[]{Keys.V}, WidthMultiplier=1 },
                new KeyVisual { Label="B", MatchCodes=new[]{Keys.B}, WidthMultiplier=1 },
                new KeyVisual { Label="N", MatchCodes=new[]{Keys.N}, WidthMultiplier=1 },
                new KeyVisual { Label="M", MatchCodes=new[]{Keys.M}, WidthMultiplier=1 },
                new KeyVisual { Label=",", MatchCodes=new[]{Keys.Oemcomma}, WidthMultiplier=1 },
                new KeyVisual { Label=".", MatchCodes=new[]{Keys.OemPeriod}, WidthMultiplier=1 },
                new KeyVisual { Label="/", MatchCodes=new[]{Keys.OemQuestion}, WidthMultiplier=1 },
                new KeyVisual { Label="Shift", MatchCodes=new[]{Keys.RShiftKey}, WidthMultiplier=2.8f }
            });

            // Row 5: Modifiers – use ONLY L/R specific codes to avoid both highlighting
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Ctrl", MatchCodes=new[]{Keys.LControlKey}, WidthMultiplier=1.5f },
                new KeyVisual { Label="Win", MatchCodes=new[]{Keys.LWin}, WidthMultiplier=1.2f },
                new KeyVisual { Label="Alt", MatchCodes=new[]{Keys.LMenu}, WidthMultiplier=1.2f },
                new KeyVisual { Label="Space", MatchCodes=new[]{Keys.Space}, WidthMultiplier=6.0f },
                new KeyVisual { Label="Alt", MatchCodes=new[]{Keys.RMenu}, WidthMultiplier=1.2f },
                new KeyVisual { Label="Win", MatchCodes=new[]{Keys.RWin}, WidthMultiplier=1.2f },
                new KeyVisual { Label="Menu", MatchCodes=new[]{Keys.Apps}, WidthMultiplier=1.2f },
                new KeyVisual { Label="Ctrl", MatchCodes=new[]{Keys.RControlKey}, WidthMultiplier=1.5f }
            });

            // --- Secondary Clusters (Nav and Arrows) ---
            // Row NavUpper
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Ins", MatchCodes=new[]{Keys.Insert}, WidthMultiplier=1 },
                new KeyVisual { Label="Home", MatchCodes=new[]{Keys.Home}, WidthMultiplier=1 },
                new KeyVisual { Label="PgUp", MatchCodes=new[]{Keys.PageUp, Keys.Prior}, WidthMultiplier=1 }
            });
            // Row NavLower
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Del", MatchCodes=new[]{Keys.Delete}, WidthMultiplier=1 },
                new KeyVisual { Label="End", MatchCodes=new[]{Keys.End}, WidthMultiplier=1 },
                new KeyVisual { Label="PgDn", MatchCodes=new[]{Keys.PageDown, Keys.Next}, WidthMultiplier=1 }
            });
            // Row ArrowsUpper
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="↑", MatchCodes=new[]{Keys.Up}, WidthMultiplier=1 }
            });
            // Row ArrowsLower
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="←", MatchCodes=new[]{Keys.Left}, WidthMultiplier=1 },
                new KeyVisual { Label="↓", MatchCodes=new[]{Keys.Down}, WidthMultiplier=1 },
                new KeyVisual { Label="→", MatchCodes=new[]{Keys.Right}, WidthMultiplier=1 }
            });
        }

        private void RenderKeyboard(Panel container)
        {
            int baseSize = 45;
            int spacing = 5;
            int maxMainX = 0;

            // Render Standard Rows (0-5)
            for (int r = 0; r < 6; r++)
            {
                int currX = 5;
                for (int c = 0; c < layoutRows[r].Count; c++)
                {
                    KeyVisual kv = layoutRows[r][c];
                    if (kv.WidthMultiplier == 0) continue;

                    int w = (int)(baseSize * kv.WidthMultiplier);
                    if (kv.MatchCodes.Length == 0) { // Gap
                        currX += w + spacing;
                        continue;
                    }

                    Panel p = new Panel() {
                        Location = new Point(currX, 5 + r * (baseSize + spacing)),
                        Size = new Size(w, baseSize),
                        BackColor = Color.FromArgb(40, 40, 45), // Modern key color
                        BorderStyle = BorderStyle.None
                    };

                    Label l = new Label() {
                        Text = kv.Label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter,
                        ForeColor = Color.LightGray, Font = new Font("Segoe UI", 9, FontStyle.Bold)
                    };
                    p.Controls.Add(l);
                    kv.Panel = p;
                    container.Controls.Add(p);
                    currX += w + spacing;
                }
                if (currX > maxMainX) maxMainX = currX;
            }

            // Dynamic X position for side clusters (prevents cutoffs)
            int sideStartX = Math.Max(maxMainX + 20, 850);
            
            // Render Navigation Cluster (Rows 6-7) – relative to sideStartX
            int navStartY = 5 + (baseSize + spacing); 
            for (int r = 6; r < 8; r++)
            {
                int currX = sideStartX;
                for (int c = 0; c < layoutRows[r].Count; c++)
                {
                    KeyVisual kv = layoutRows[r][c];
                    int w = (int)(baseSize * kv.WidthMultiplier);
                    Panel p = CreateKeyPanel(kv, currX, navStartY + (r-6) * (baseSize + spacing), w, baseSize);
                    container.Controls.Add(p);
                    currX += w + spacing;
                }
            }

            // Render Arrow Cluster (Rows 8-9)
            int arrowStartY = 5 + 4 * (baseSize + spacing);
            KeyVisual up = layoutRows[8][0];
            container.Controls.Add(CreateKeyPanel(up, sideStartX + baseSize + spacing, arrowStartY, baseSize, baseSize));
            
            int currArrowX = sideStartX;
            for (int c = 0; c < layoutRows[9].Count; c++)
            {
                KeyVisual kv = layoutRows[9][c];
                container.Controls.Add(CreateKeyPanel(kv, currArrowX, arrowStartY + baseSize + spacing, baseSize, baseSize));
                currArrowX += baseSize + spacing;
            }
        }

        private Panel CreateKeyPanel(KeyVisual kv, int x, int y, int w, int h)
        {
            Panel p = new Panel() {
                Location = new Point(x, y),
                Size = new Size(w, h),
                BackColor = Color.FromArgb(40, 40, 45),
                BorderStyle = BorderStyle.None
            };
            Label l = new Label() {
                Text = kv.Label, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter,
                ForeColor = Color.LightGray, Font = new Font("Segoe UI", 9, FontStyle.Bold)
            };
            p.Controls.Add(l);
            kv.Panel = p;
            return p;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern short GetAsyncKeyState(int vKey);

        private void QaKeyboardForm_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;
            e.SuppressKeyPress = true;

            // Debug mapping for special keys
            Keys key = e.KeyCode;
            
            // WinForms does not distinguish left and right modifiers in KeyEventArgs natively.
            // We use GetAsyncKeyState to check their physical states.
            if (key == Keys.ShiftKey)
            {
                if ((GetAsyncKeyState((int)Keys.LShiftKey) & 0x8000) != 0) key = Keys.LShiftKey;
                else if ((GetAsyncKeyState((int)Keys.RShiftKey) & 0x8000) != 0) key = Keys.RShiftKey;
            }
            else if (key == Keys.ControlKey)
            {
                if ((GetAsyncKeyState((int)Keys.LControlKey) & 0x8000) != 0) key = Keys.LControlKey;
                else if ((GetAsyncKeyState((int)Keys.RControlKey) & 0x8000) != 0) key = Keys.RControlKey;
            }
            else if (key == Keys.Menu || key == Keys.Alt)
            {
                if ((GetAsyncKeyState((int)Keys.LMenu) & 0x8000) != 0) key = Keys.LMenu;
                else if ((GetAsyncKeyState((int)Keys.RMenu) & 0x8000) != 0) key = Keys.RMenu;
            }

            // Find matching key visual
            KeyVisual found = null;
            foreach (var row in layoutRows)
            {
                foreach (var kv in row)
                {
                    if (kv.MatchCodes.Contains(key))
                    {
                        found = kv;
                        break;
                    }
                }
                if (found != null) break;
            }

            if (found != null)
            {
                found.Clicks++;
                Color newColor = Color.FromArgb(60, 60, 60);
                if (found.Clicks == 1) newColor = Color.LimeGreen;
                else if (found.Clicks == 2) newColor = Color.MediumOrchid;
                else if (found.Clicks >= 3) newColor = Color.DeepSkyBlue;

                found.Panel.BackColor = newColor;
                Label l = (Label)found.Panel.Controls[0];
                l.ForeColor = Color.White;
            }
        }
    }
}
