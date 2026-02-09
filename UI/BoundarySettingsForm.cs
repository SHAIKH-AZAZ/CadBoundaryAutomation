using System;
using WinForms = System.Windows.Forms;
using Drawing = System.Drawing;

namespace CadBoundaryAutomation.UI
{
    public enum BarsOrientation
    {
        Horizontal,
        Vertical,
        Both
    }

    public class BoundarySettingsForm : WinForms.Form
    {
        private readonly WinForms.NumericUpDown nudSpacingH;
        private readonly WinForms.NumericUpDown nudSpacingV;
        private readonly WinForms.ComboBox cmbOrientation;

        public double SpacingH => (double)nudSpacingH.Value;
        public double SpacingV => (double)nudSpacingV.Value;

        public BarsOrientation Orientation
        {
            get
            {
                switch (cmbOrientation.SelectedIndex)
                {
                    case 1: return BarsOrientation.Vertical;
                    case 2: return BarsOrientation.Both;
                    default: return BarsOrientation.Horizontal;
                }
            }
        }

        public BoundarySettingsForm(double defaultSpacingH, double defaultSpacingV, BarsOrientation defaultOrientation)
        {
            Text = "Boundary Settings";
            FormBorderStyle = WinForms.FormBorderStyle.FixedDialog;
            StartPosition = WinForms.FormStartPosition.CenterScreen;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Drawing.Size(420, 210);

            var lblOri = new WinForms.Label
            {
                Text = "Orientation:",
                AutoSize = true,
                Location = new Drawing.Point(20, 20)
            };

            cmbOrientation = new WinForms.ComboBox
            {
                Location = new Drawing.Point(200, 16),
                Width = 180,
                DropDownStyle = WinForms.ComboBoxStyle.DropDownList
            };
            cmbOrientation.Items.AddRange(new object[] { "Horizontal", "Vertical", "Both" });
            cmbOrientation.SelectedIndex =
                defaultOrientation == BarsOrientation.Vertical ? 1 :
                defaultOrientation == BarsOrientation.Both ? 2 : 0;

            cmbOrientation.SelectedIndexChanged += (s, e) => UpdateEnableState();

            var lblH = new WinForms.Label
            {
                Text = "Horizontal spacing (mm):",
                AutoSize = true,
                Location = new Drawing.Point(20, 65)
            };

            nudSpacingH = new WinForms.NumericUpDown
            {
                Location = new Drawing.Point(200, 61),
                Width = 180,
                DecimalPlaces = 2,
                Minimum = 0.01m,
                Maximum = 1000000m,
                Increment = 10m,
                Value = (decimal)Math.Max(0.01, defaultSpacingH)
            };

            var lblV = new WinForms.Label
            {
                Text = "Vertical spacing (mm):",
                AutoSize = true,
                Location = new Drawing.Point(20, 105)
            };

            nudSpacingV = new WinForms.NumericUpDown
            {
                Location = new Drawing.Point(200, 101),
                Width = 180,
                DecimalPlaces = 2,
                Minimum = 0.01m,
                Maximum = 1000000m,
                Increment = 10m,
                Value = (decimal)Math.Max(0.01, defaultSpacingV)
            };

            var btnOk = new WinForms.Button
            {
                Text = "OK",
                DialogResult = WinForms.DialogResult.OK,
                Location = new Drawing.Point(225, 150),
                Width = 75
            };

            var btnCancel = new WinForms.Button
            {
                Text = "Cancel",
                DialogResult = WinForms.DialogResult.Cancel,
                Location = new Drawing.Point(305, 150),
                Width = 75
            };

            AcceptButton = btnOk;
            CancelButton = btnCancel;

            Controls.Add(lblOri);
            Controls.Add(cmbOrientation);
            Controls.Add(lblH);
            Controls.Add(nudSpacingH);
            Controls.Add(lblV);
            Controls.Add(nudSpacingV);
            Controls.Add(btnOk);
            Controls.Add(btnCancel);

            UpdateEnableState();
        }

        private void UpdateEnableState()
        {
            var ori = Orientation;

            if (ori == BarsOrientation.Horizontal)
            {
                nudSpacingH.Enabled = true;
                nudSpacingV.Enabled = false;
            }
            else if (ori == BarsOrientation.Vertical)
            {
                nudSpacingH.Enabled = false;
                nudSpacingV.Enabled = true;
            }
            else
            {
                nudSpacingH.Enabled = true;
                nudSpacingV.Enabled = true;
            }
        }
    }
}
