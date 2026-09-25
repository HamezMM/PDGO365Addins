using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace SpeakerRename.UI
{
    /// <summary>
    /// Wizard-style dialog: one unique speaker at a time (Back / Next), with a list of all
    /// speakers on the left for jumping around. Names are written straight onto the
    /// <see cref="SpeakerField"/> objects; the caller replaces text only on
    /// <see cref="DialogResult.OK"/>. Built in code (no designer file) to match the rest of
    /// the repo's hand-maintained projects.
    /// </summary>
    internal sealed class RenameSpeakersForm : Form
    {
        private readonly IReadOnlyList<SpeakerField> _fields;
        private int _current = -1;

        private readonly ListBox _fieldList = new ListBox();
        private readonly Label _progressLabel = new Label();
        private readonly Label _speakerLabel = new Label();
        private readonly Label _occurrencesLabel = new Label();
        private readonly Label _contextLabel = new Label();
        private readonly TextBox _valueTextBox = new TextBox();
        private readonly Label _filledLabel = new Label();
        private readonly Button _backButton = new Button();
        private readonly Button _nextButton = new Button();
        private readonly Button _replaceButton = new Button();
        private readonly Button _cancelButton = new Button();

        public RenameSpeakersForm(IReadOnlyList<SpeakerField> fields, string documentName)
        {
            _fields = fields ?? throw new ArgumentNullException(nameof(fields));
            BuildLayout(documentName);

            foreach (SpeakerField field in _fields) _fieldList.Items.Add(field);
            UpdateFilledLabel();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (_fields.Count > 0) _fieldList.SelectedIndex = 0;
        }

        // ---- navigation ------------------------------------------------------------

        private void ShowField(int index)
        {
            if (index < 0 || index >= _fields.Count) return;
            _current = index;
            SpeakerField field = _fields[index];

            _progressLabel.Text = $"Speaker {index + 1} of {_fields.Count}";
            _speakerLabel.Text = field.Label;
            _occurrencesLabel.Text = field.Occurrences == 1
                ? "Appears once in the document."
                : $"Appears {field.Occurrences} times in the document.";
            _contextLabel.Text = field.Context;

            // Assigning Text raises TextChanged; _current already points at this field,
            // so that just writes the same value back.
            _valueTextBox.Text = field.Value ?? string.Empty;
            _valueTextBox.SelectAll();
            _valueTextBox.Focus();

            bool isLast = index == _fields.Count - 1;
            _backButton.Enabled = index > 0;
            _nextButton.Text = isLast ? "&Finish" : "&Next >";
        }

        private void FieldList_SelectedIndexChanged(object sender, EventArgs e) => ShowField(_fieldList.SelectedIndex);

        private void BackButton_Click(object sender, EventArgs e)
        {
            if (_current > 0) _fieldList.SelectedIndex = _current - 1;
        }

        private void NextButton_Click(object sender, EventArgs e)
        {
            if (_current < _fields.Count - 1) _fieldList.SelectedIndex = _current + 1;
            else TryFinish();
        }

        private void ReplaceButton_Click(object sender, EventArgs e) => TryFinish();

        private void TryFinish()
        {
            if (!_fields.Any(f => f.HasValue))
            {
                MessageBox.Show(this, "Enter a name for at least one speaker, or click Cancel.",
                    Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
                _valueTextBox.Focus();
                return;
            }

            DialogResult = DialogResult.OK;
            Close();
        }

        private void ValueTextBox_TextChanged(object sender, EventArgs e)
        {
            if (_current < 0) return;
            _fields[_current].Value = _valueTextBox.Text;
            _fieldList.Invalidate(_fieldList.GetItemRectangle(_current));
            UpdateFilledLabel();
        }

        private void UpdateFilledLabel()
        {
            int filled = _fields.Count(f => f.HasValue);
            _filledLabel.Text = $"{filled} of {_fields.Count} named";
        }

        private void FieldList_DrawItem(object sender, DrawItemEventArgs e)
        {
            e.DrawBackground();
            if (e.Index >= 0)
            {
                TextRenderer.DrawText(e.Graphics, _fields[e.Index].ToString(), e.Font, e.Bounds, e.ForeColor,
                    TextFormatFlags.Left | TextFormatFlags.VerticalCenter | TextFormatFlags.EndEllipsis | TextFormatFlags.NoPrefix);
            }
            e.DrawFocusRectangle();
        }

        // ---- layout ----------------------------------------------------------------

        private void BuildLayout(string documentName)
        {
            SuspendLayout();

            Text = "Rename Speakers — " + documentName;
            Font = new Font("Segoe UI", 9F);
            AutoScaleDimensions = new SizeF(96F, 96F);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(680, 360);
            MinimumSize = new Size(560, 340);
            FormBorderStyle = FormBorderStyle.Sizable;
            StartPosition = FormStartPosition.CenterParent;
            ShowInTaskbar = false;
            ShowIcon = false;
            MinimizeBox = false;
            MaximizeBox = false;

            // Left: every unique speaker, ticked once it has a name.
            _fieldList.Dock = DockStyle.Fill;
            _fieldList.IntegralHeight = false;
            _fieldList.DrawMode = DrawMode.OwnerDrawFixed;
            _fieldList.ItemHeight = 22;
            _fieldList.DrawItem += FieldList_DrawItem;
            _fieldList.SelectedIndexChanged += FieldList_SelectedIndexChanged;

            // Right: the current speaker.
            _progressLabel.AutoSize = true;
            _progressLabel.ForeColor = SystemColors.GrayText;

            _speakerLabel.AutoSize = true;
            _speakerLabel.Font = new Font("Segoe UI", 13F, FontStyle.Bold);
            _speakerLabel.UseMnemonic = false;
            _speakerLabel.Margin = new Padding(3, 4, 3, 0);

            _occurrencesLabel.AutoSize = true;

            _contextLabel.Dock = DockStyle.Fill;
            _contextLabel.Height = 44;
            _contextLabel.AutoEllipsis = true;
            _contextLabel.UseMnemonic = false;
            _contextLabel.ForeColor = SystemColors.GrayText;
            _contextLabel.Font = new Font("Segoe UI", 9F, FontStyle.Italic);
            _contextLabel.Margin = new Padding(3, 8, 3, 8);

            var renameToLabel = new Label { Text = "&Rename to:", AutoSize = true };

            _valueTextBox.Dock = DockStyle.Fill;
            _valueTextBox.TextChanged += ValueTextBox_TextChanged;

            var hintLabel = new Label
            {
                Text = "Every occurrence of this speaker's label is replaced with the name you enter. Leave blank to skip it.",
                AutoSize = true,
                ForeColor = SystemColors.GrayText,
            };

            var detail = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                Padding = new Padding(12, 0, 0, 0),
            };
            detail.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            detail.Controls.Add(_progressLabel);
            detail.Controls.Add(_speakerLabel);
            detail.Controls.Add(_occurrencesLabel);
            detail.Controls.Add(_contextLabel);
            detail.Controls.Add(renameToLabel);
            detail.Controls.Add(_valueTextBox);
            detail.Controls.Add(hintLabel);
            for (int i = 0; i < detail.Controls.Count; i++)
            {
                detail.RowStyles.Add(detail.Controls[i] == _contextLabel
                    ? new RowStyle(SizeType.Absolute, 60F) // room for ~2 lines of snippet
                    : new RowStyle(SizeType.AutoSize));
            }
            detail.RowCount = detail.Controls.Count + 1;
            detail.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));

            // Bottom: status on the left, buttons on the right.
            _filledLabel.AutoSize = true;
            _filledLabel.Anchor = AnchorStyles.Left;
            _filledLabel.ForeColor = SystemColors.GrayText;

            ConfigureButton(_backButton, "< &Back", BackButton_Click);
            ConfigureButton(_nextButton, "&Next >", NextButton_Click);
            ConfigureButton(_replaceButton, "Rename &All", ReplaceButton_Click);
            ConfigureButton(_cancelButton, "Cancel", null);
            _cancelButton.DialogResult = DialogResult.Cancel;
            _replaceButton.Margin = new Padding(16, 3, 3, 3);

            var buttons = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                Anchor = AnchorStyles.Right,
            };
            buttons.Controls.AddRange(new Control[] { _backButton, _nextButton, _replaceButton, _cancelButton });

            var root = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 2,
                Padding = new Padding(12),
            };
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 210F));
            root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.Percent, 100F));
            root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            root.Controls.Add(_fieldList, 0, 0);
            root.Controls.Add(detail, 1, 0);
            root.Controls.Add(_filledLabel, 0, 1);
            root.Controls.Add(buttons, 1, 1);

            Controls.Add(root);

            // Enter = Next (Finish on the last speaker); Esc = Cancel.
            AcceptButton = _nextButton;
            CancelButton = _cancelButton;

            ResumeLayout(false);
            PerformLayout();
        }

        private static void ConfigureButton(Button button, string text, EventHandler onClick)
        {
            button.Text = text;
            button.AutoSize = true;
            button.MinimumSize = new Size(88, 28);
            button.UseVisualStyleBackColor = true;
            if (onClick != null) button.Click += onClick;
        }
    }
}
