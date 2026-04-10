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
        private Label lblCount;

        public QaKeyboardForm()
        {
            this.Text = "Hardware Check: Keyboard Test (מבדק מקלדת)";
            this.Size = new Size(950, 500);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.KeyPreview = true;
            this.BackColor = Color.FromArgb(30, 30, 30);
            this.RightToLeft = RightToLeft.No; // Ensure LTR for QWERTY

            Label lblInstr = new Label() { 
                Text = "הקש על מקשי המקלדת. צבע המקש ישתנה בהתאם למספר הלחיצות:\n" +
                       "אפור (לא נלחץ) ➔ ירוק (פעם 1) ➔ סגול (פעם 2) ➔ תכלת (פעם 3)\n" +
                       "שים לב: הלייאוט מציג מקלדת אמריקאית סטנדרטית.", 
                Dock = DockStyle.Top, Height = 90, ForeColor = Color.White, 
                Font = new Font("Segoe UI", 12, FontStyle.Regular), TextAlign = ContentAlignment.MiddleCenter,
                RightToLeft = RightToLeft.Yes
            };
            this.Controls.Add(lblInstr);

            Panel keyboardContainer = new Panel() { 
                Location = new Point(20, 100), Size = new Size(900, 300) 
            };
            this.Controls.Add(keyboardContainer);

            BuildKeyboardLayout();
            RenderKeyboard(keyboardContainer);

            Button btnDone = new Button() { 
                Text = "סיימתי את הבדיקה", Dock = DockStyle.Bottom, Height = 55, 
                BackColor = Color.SeaGreen, ForeColor = Color.White, 
                Font = new Font("Segoe UI", 12, FontStyle.Bold), Cursor = Cursors.Hand 
            };
            btnDone.FlatAppearance.BorderSize = 0;
            btnDone.FlatStyle = FlatStyle.Flat;
            btnDone.Click += (s, e) => this.Close();
            this.Controls.Add(btnDone);

            this.KeyDown += QaKeyboardForm_KeyDown;
        }

        private void BuildKeyboardLayout()
        {
            layoutRows = new List<List<KeyVisual>>();

            // Row 1
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

            // Row 2
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

            // Row 3
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

            // Row 4
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Shift", MatchCodes=new[]{Keys.LShiftKey, Keys.ShiftKey, Keys.Shift}, WidthMultiplier=2.2f },
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

            // Row 5
            layoutRows.Add(new List<KeyVisual> {
                new KeyVisual { Label="Ctrl", MatchCodes=new[]{Keys.LControlKey, Keys.ControlKey, Keys.Control}, WidthMultiplier=1.5f },
                new KeyVisual { Label="Win", MatchCodes=new[]{Keys.LWin}, WidthMultiplier=1.5f },
                new KeyVisual { Label="Alt", MatchCodes=new[]{Keys.LMenu, Keys.Alt, Keys.Menu}, WidthMultiplier=1.5f },
                new KeyVisual { Label="", MatchCodes=new[]{Keys.Space}, WidthMultiplier=6.0f },
                new KeyVisual { Label="Alt", MatchCodes=new[]{Keys.RMenu}, WidthMultiplier=1.5f },
                new KeyVisual { Label="Win", MatchCodes=new[]{Keys.RWin}, WidthMultiplier=1.2f },
                new KeyVisual { Label="Menu", MatchCodes=new[]{Keys.Apps}, WidthMultiplier=1.2f },
                new KeyVisual { Label="Ctrl", MatchCodes=new[]{Keys.RControlKey}, WidthMultiplier=1.6f } // Slightly thicker
            });
        }

        private void RenderKeyboard(Panel container)
        {
            int startX = 10;
            int startY = 10;
            int baseSize = 50;
            int spacing = 5;

            for (int r = 0; r < layoutRows.Count; r++)
            {
                int currX = startX;
                for (int c = 0; c < layoutRows[r].Count; c++)
                {
                    KeyVisual kv = layoutRows[r][c];
                    
                    int w = (int)(baseSize * kv.WidthMultiplier);
                    Panel p = new Panel() {
                        Location = new Point(currX, startY + r * (baseSize + spacing)),
                        Size = new Size(w, baseSize),
                        BackColor = Color.FromArgb(60, 60, 60),
                        BorderStyle = BorderStyle.FixedSingle
                    };

                    Label l = new Label() {
                        Text = kv.Label,
                        Dock = DockStyle.Fill,
                        TextAlign = ContentAlignment.MiddleCenter,
                        ForeColor = Color.LightGray,
                        Font = new Font("Segoe UI", 10, FontStyle.Bold)
                    };
                    p.Controls.Add(l);
                    kv.Panel = p;
                    container.Controls.Add(p);

                    currX += w + spacing;
                }
            }
        }

        private void QaKeyboardForm_KeyDown(object sender, KeyEventArgs e)
        {
            e.Handled = true;

            // Find matching key visual
            KeyVisual found = null;
            foreach (var row in layoutRows)
            {
                foreach (var kv in row)
                {
                    if (kv.MatchCodes.Contains(e.KeyCode))
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
                // Gray -> Green -> Purple -> Cyan
                Color newColor = Color.FromArgb(60, 60, 60);
                if (found.Clicks == 1) newColor = Color.LimeGreen;
                else if (found.Clicks == 2) newColor = Color.MediumOrchid;
                else if (found.Clicks >= 3) newColor = Color.DeepSkyBlue;

                found.Panel.BackColor = newColor;

                // Visual blip effect
                Label l = (Label)found.Panel.Controls[0];
                l.ForeColor = Color.White;
            }
        }
    }
}
