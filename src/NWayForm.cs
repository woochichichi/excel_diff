using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelDiffMerge
{
    /// <summary>
    /// N-way 취합 뷰 (spec §5-5, §6).
    ///  - base 1개 + 나머지 버전 파일들을 선택 → base 대비 diff.
    ///  - 셀별로 "어느 버전이 뭘 바꿨나" 표(그리드)로 표시. 충돌은 주황.
    ///  - 셀 클릭 → 각 버전 값 중 채택할 값 선택 → base 파일에 병합/저장.
    /// 1차: 조회 + 충돌 강조 + 단일 값 채택 저장. 상세 인터랙션은 실사용 피드백으로(§11).
    /// </summary>
    public sealed class NWayForm : Form
    {
        private static readonly Color ColConflict = Color.FromArgb(255, 200, 140); // 주황
        private static readonly Color ColChanged = Color.FromArgb(255, 245, 170);  // 노랑

        private readonly Settings _settings;

        private TextBox _txtBase;
        private ListBox _lstVersions;
        private TabControl _tabs;
        private DataGridView _grid;
        private ToolStripStatusLabel _lblSummary;
        private ToolStripProgressBar _progress;
        private BackgroundWorker _worker;

        private string _basePath;
        private readonly List<string> _versionPaths = new List<string>();
        private NWayResult _result;
        private NWaySheetDiff _curSheet;

        // 병합 채택: base 에 적용할 값. key "sheet:row:col" → value.
        private readonly Dictionary<string, object> _adopt = new Dictionary<string, object>();

        public NWayForm(Settings settings)
        {
            _settings = settings;
            BuildUi();
        }

        private void BuildUi()
        {
            Text = "N-way 취합 (base 대비 다버전 비교)";
            Width = 1180; Height = 760;
            StartPosition = FormStartPosition.CenterParent;

            ToolStrip tool = new ToolStrip();
            tool.GripStyle = ToolStripGripStyle.Hidden;
            AddBtn(tool, "base 선택", delegate { PickBase(); });
            AddBtn(tool, "버전 추가", delegate { AddVersion(); });
            AddBtn(tool, "버전 제거", delegate { RemoveVersion(); });
            tool.Items.Add(new ToolStripSeparator());
            AddBtn(tool, "비교", delegate { StartCompare(); });
            tool.Items.Add(new ToolStripSeparator());
            AddBtn(tool, "선택 셀 값 채택…", delegate { AdoptSelected(); });
            AddBtn(tool, "base 에 저장", delegate { SaveAdopted(); });
            Controls.Add(tool);
            tool.Dock = DockStyle.Top;

            // 상단: base + 버전 목록
            Panel top = new Panel();
            top.Dock = DockStyle.Top;
            top.Height = 120;

            Label lb = new Label();
            lb.Text = "base:";
            lb.Location = new Point(8, 8);
            lb.AutoSize = true;
            top.Controls.Add(lb);

            _txtBase = new TextBox();
            _txtBase.Location = new Point(60, 4);
            _txtBase.Width = 1060;
            _txtBase.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            _txtBase.ReadOnly = true;
            top.Controls.Add(_txtBase);

            Label lv = new Label();
            lv.Text = "버전들:";
            lv.Location = new Point(8, 34);
            lv.AutoSize = true;
            top.Controls.Add(lv);

            _lstVersions = new ListBox();
            _lstVersions.Location = new Point(60, 32);
            _lstVersions.Width = 1060;
            _lstVersions.Height = 80;
            _lstVersions.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            top.Controls.Add(_lstVersions);

            Controls.Add(top);

            _tabs = new TabControl();
            _tabs.Dock = DockStyle.Top;
            _tabs.Height = 28;
            _tabs.SelectedIndexChanged += delegate { OnSheetChanged(); };
            Controls.Add(_tabs);

            _grid = new DataGridView();
            _grid.Dock = DockStyle.Fill;
            _grid.ReadOnly = true;
            _grid.AllowUserToAddRows = false;
            _grid.RowHeadersVisible = true;
            _grid.SelectionMode = DataGridViewSelectionMode.CellSelect;
            _grid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            _grid.CellFormatting += Grid_CellFormatting;
            Controls.Add(_grid);
            _grid.BringToFront();

            StatusStrip status = new StatusStrip();
            _lblSummary = new ToolStripStatusLabel("base 와 버전들을 지정하고 [비교]를 누르세요.");
            _lblSummary.Spring = true;
            _lblSummary.TextAlign = ContentAlignment.MiddleLeft;
            _progress = new ToolStripProgressBar();
            _progress.Style = ProgressBarStyle.Marquee;
            _progress.Visible = false;
            status.Items.Add(_lblSummary);
            status.Items.Add(_progress);
            Controls.Add(status);

            _worker = new BackgroundWorker();
            _worker.DoWork += Worker_DoWork;
            _worker.RunWorkerCompleted += Worker_Completed;
        }

        private static void AddBtn(ToolStrip t, string text, EventHandler h)
        {
            ToolStripButton b = new ToolStripButton(text);
            b.DisplayStyle = ToolStripItemDisplayStyle.Text;
            b.Click += h;
            t.Items.Add(b);
        }

        private void PickBase()
        {
            string f = PickFile();
            if (f != null) { _basePath = f; _txtBase.Text = f; }
        }

        private void AddVersion()
        {
            string f = PickFile();
            if (f != null) { _versionPaths.Add(f); _lstVersions.Items.Add(f); }
        }

        private void RemoveVersion()
        {
            int i = _lstVersions.SelectedIndex;
            if (i < 0) return;
            _versionPaths.RemoveAt(i);
            _lstVersions.Items.RemoveAt(i);
        }

        private string PickFile()
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel 파일 (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|모든 파일 (*.*)|*.*";
                string last = _settings.LastFolder;
                if (!string.IsNullOrEmpty(last) && Directory.Exists(last))
                    dlg.InitialDirectory = last;
                if (dlg.ShowDialog(this) != DialogResult.OK) return null;
                _settings.LastFolder = Path.GetDirectoryName(dlg.FileName);
                _settings.Save();
                return dlg.FileName;
            }
        }

        private void StartCompare()
        {
            if (string.IsNullOrEmpty(_basePath) || _versionPaths.Count == 0)
            {
                MessageBox.Show(this, "base 파일과 버전 1개 이상을 지정하세요.", "안내",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_worker.IsBusy) return;
            _adopt.Clear();
            _progress.Visible = true;
            _lblSummary.Text = "로드/비교 중…";
            _worker.RunWorkerAsync();
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            using (ExcelComReader reader = new ExcelComReader())
            {
                WorkbookData baseWb = reader.LoadWorkbook(_basePath);
                List<WorkbookData> versions = new List<WorkbookData>();
                foreach (string p in _versionPaths)
                    versions.Add(reader.LoadWorkbook(p));
                e.Result = NWayDiffEngine.Compare(baseWb, versions);
            }
        }

        private void Worker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            _progress.Visible = false;
            if (e.Error != null)
            {
                _lblSummary.Text = "오류: " + e.Error.Message;
                MessageBox.Show(this, e.Error.Message, "비교 실패",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            _result = (NWayResult)e.Result;
            PopulateTabs();
            _lblSummary.Text = _result.Summary();
        }

        private void PopulateTabs()
        {
            _tabs.TabPages.Clear();
            foreach (NWaySheetDiff s in _result.Sheets)
            {
                string marker = (s.ChangedCount > 0) ? " ●" : "";
                TabPage tp = new TabPage(s.Name + marker);
                tp.Tag = s;
                _tabs.TabPages.Add(tp);
            }
            if (_tabs.TabPages.Count > 0) { _tabs.SelectedIndex = 0; OnSheetChanged(); }
        }

        private void OnSheetChanged()
        {
            if (_tabs.SelectedTab == null) return;
            _curSheet = (NWaySheetDiff)_tabs.SelectedTab.Tag;
            BuildGrid();
        }

        private void BuildGrid()
        {
            _grid.Rows.Clear();
            _grid.Columns.Clear();
            if (_curSheet == null) return;

            _grid.Columns.Add("loc", "위치");
            _grid.Columns.Add("base", "base");
            for (int i = 0; i < _result.VersionPaths.Count; i++)
                _grid.Columns.Add("v" + i, "v" + (i + 1) + ": " + Path.GetFileName(_result.VersionPaths[i]));
            _grid.Columns.Add("state", "상태");

            foreach (NWayCell cell in _curSheet.Cells)
            {
                object[] rowVals = new object[_grid.Columns.Count];
                rowVals[0] = MainForm.ColLetter(cell.Col) + cell.Row;
                rowVals[1] = DiffEngine.ToText(cell.BaseValue);
                for (int i = 0; i < _result.VersionPaths.Count; i++)
                {
                    object v;
                    if (cell.Changes.TryGetValue(i, out v))
                        rowVals[2 + i] = DiffEngine.ToText(v);
                    else
                        rowVals[2 + i] = ""; // base 와 동일
                }
                rowVals[rowVals.Length - 1] = cell.Conflict ? "충돌" : "변경";
                int idx = _grid.Rows.Add(rowVals);
                _grid.Rows[idx].Tag = cell;
            }
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _grid.Rows.Count) return;
            NWayCell cell = _grid.Rows[e.RowIndex].Tag as NWayCell;
            if (cell == null) return;
            Color c = cell.Conflict ? ColConflict : ColChanged;
            e.CellStyle.BackColor = c;
        }

        /// <summary>선택 행의 셀에 대해 채택할 값을 고른다(base 또는 각 버전 값).</summary>
        private void AdoptSelected()
        {
            if (_curSheet == null || _grid.CurrentRow == null) return;
            NWayCell cell = _grid.CurrentRow.Tag as NWayCell;
            if (cell == null) return;

            // 후보값 수집(중복 제거).
            List<object> candidates = new List<object>();
            List<string> labels = new List<string>();
            candidates.Add(cell.BaseValue); labels.Add("base: " + DiffEngine.ToText(cell.BaseValue));
            foreach (KeyValuePair<int, object> kv in cell.Changes)
            {
                candidates.Add(kv.Value);
                labels.Add("v" + (kv.Key + 1) + ": " + DiffEngine.ToText(kv.Value));
            }

            using (Form dlg = new Form())
            {
                dlg.Text = "채택할 값 선택 — " + MainForm.ColLetter(cell.Col) + cell.Row;
                dlg.Width = 460; dlg.Height = 260;
                dlg.StartPosition = FormStartPosition.CenterParent;
                dlg.FormBorderStyle = FormBorderStyle.FixedDialog;
                dlg.MaximizeBox = false; dlg.MinimizeBox = false;

                ListBox lb = new ListBox();
                lb.Dock = DockStyle.Fill;
                foreach (string s in labels) lb.Items.Add(s);
                lb.SelectedIndex = 0;
                dlg.Controls.Add(lb);

                Button ok = new Button();
                ok.Text = "채택";
                ok.Dock = DockStyle.Bottom;
                ok.DialogResult = DialogResult.OK;
                dlg.Controls.Add(ok);
                dlg.AcceptButton = ok;

                if (dlg.ShowDialog(this) == DialogResult.OK && lb.SelectedIndex >= 0)
                {
                    string key = _curSheet.Name + ":" + cell.Row + ":" + cell.Col;
                    _adopt[key] = candidates[lb.SelectedIndex];
                    _lblSummary.Text = string.Format("{0}  |  채택 대기 {1}개", _result.Summary(), _adopt.Count);
                }
            }
        }

        private void SaveAdopted()
        {
            if (_adopt.Count == 0)
            {
                MessageBox.Show(this, "채택한 값이 없습니다.", "안내",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            bool backup = MessageBox.Show(this, "저장 전 백업 복사본을 만들까요? (권장: 예)",
                "백업", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

            List<MergeItem> items = new List<MergeItem>();
            foreach (KeyValuePair<string, object> kv in _adopt)
            {
                int lastColon = kv.Key.LastIndexOf(':');
                int prevColon = kv.Key.LastIndexOf(':', lastColon - 1);
                string sheet = kv.Key.Substring(0, prevColon);
                int row = int.Parse(kv.Key.Substring(prevColon + 1, lastColon - prevColon - 1));
                int col = int.Parse(kv.Key.Substring(lastColon + 1));
                items.Add(new MergeItem(sheet, row, col, kv.Value));
            }

            try
            {
                using (MergeEngine me = new MergeEngine())
                {
                    string backupPath;
                    me.ApplyAndSave(_basePath, items, backup, out backupPath);
                }
                _adopt.Clear();
                MessageBox.Show(this, "base 파일에 저장 완료.", "완료",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "저장 실패(원본 보존):\r\n" + ex.Message, "저장 실패",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
