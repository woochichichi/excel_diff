using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace ExcelDiffMerge
{
    /// <summary>
    /// WinMerge 식 좌우 병렬 뷰 (spec §6).
    ///  - DataGridView 2개(가상 모드) + 색상 하이라이트 + 동기 스크롤 + 시트 탭.
    ///  - 변경점 네비게이션(F7/F8), 요약 패널, "변경만 보기" 필터.
    ///  - 방향 병합(좌→우 / 우→좌) → 저장(백업 옵션).
    /// 대용량 대비 파일 로드/비교는 BackgroundWorker 로 수행(UI 블로킹 방지).
    /// </summary>
    public sealed class MainForm : Form
    {
        // 색상 (spec §5-3, §6)
        private static readonly Color ColChanged = Color.FromArgb(255, 245, 170); // 노랑
        private static readonly Color ColAdded = Color.FromArgb(190, 240, 190);   // 초록
        private static readonly Color ColDeleted = Color.FromArgb(250, 190, 190); // 빨강
        private static readonly Color ColMerged = Color.FromArgb(190, 215, 255);  // 병합 적용됨(파랑)

        private readonly Settings _settings;

        private ToolStrip _tool;
        private DataGridView _gridLeft;
        private DataGridView _gridRight;
        private TabControl _tabs;
        private Label _lblLeftPath;
        private Label _lblRightPath;
        private ToolStripStatusLabel _lblSummary;
        private ToolStripProgressBar _progress;
        private ToolStripButton _btnChangesOnly;
        private BackgroundWorker _worker;

        // 상태
        private string _leftPath;
        private string _rightPath;
        private DiffResult _diff;
        private SheetDiff _curSheet;
        private SheetData _curLeft;
        private SheetData _curRight;

        // 가상 그리드 매핑
        private int[] _rowMap = new int[0];   // 표시행 인덱스 → 절대행(1-based)
        private Dictionary<int, int> _absRowToDisplay = new Dictionary<int, int>();
        private int _colStart = 1;            // 표시열 0 == 절대열 _colStart
        private int _colCount;

        private bool _changesOnly;
        private bool _syncing; // 동기 스크롤/선택 재진입 방지

        // 병합 대기(각 파일에 적용할 값). key = "sheet:row:col".
        private readonly Dictionary<string, object> _pendingLeft = new Dictionary<string, object>();
        private readonly Dictionary<string, object> _pendingRight = new Dictionary<string, object>();

        // 네비게이션용 현재 시트 diff 셀(정렬).
        private List<DiffCell> _navCells = new List<DiffCell>();
        private int _navIndex = -1;

        public MainForm()
        {
            _settings = Settings.LoadDefault();
            BuildUi();
            ShowOnboardingIfNeeded();
        }

        // ---------------------------------------------------------------- UI 구성

        private void BuildUi()
        {
            Text = "Excel Diff / Merge";
            Width = 1280;
            Height = 800;
            StartPosition = FormStartPosition.CenterScreen;
            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;
            KeyPreview = true;
            KeyDown += OnKeyDown;

            // 툴바
            ToolStrip tool = new ToolStrip();
            _tool = tool;
            tool.GripStyle = ToolStripGripStyle.Hidden;
            AddButton(tool, "좌측 열기", delegate { OpenFile(true); });
            AddButton(tool, "우측 열기", delegate { OpenFile(false); });
            tool.Items.Add(new ToolStripSeparator());
            AddButton(tool, "비교", delegate { StartCompare(); });
            tool.Items.Add(new ToolStripSeparator());
            AddButton(tool, "◀ 이전차이(F8)", delegate { NavigateDiff(-1); });
            AddButton(tool, "다음차이(F7) ▶", delegate { NavigateDiff(1); });
            tool.Items.Add(new ToolStripSeparator());
            _btnChangesOnly = new ToolStripButton("변경만 보기");
            _btnChangesOnly.CheckOnClick = true;
            _btnChangesOnly.CheckedChanged += delegate
            {
                _changesOnly = _btnChangesOnly.Checked;
                RefreshGridRows();
            };
            tool.Items.Add(_btnChangesOnly);
            tool.Items.Add(new ToolStripSeparator());
            AddButton(tool, "좌→우 복사", delegate { MergeSelected(true); });
            AddButton(tool, "우→좌 복사", delegate { MergeSelected(false); });
            AddButton(tool, "저장", delegate { SaveMerges(); });
            tool.Items.Add(new ToolStripSeparator());
            AddButton(tool, "N-way…", delegate { OpenNWayDialog(); });
            tool.Dock = DockStyle.Top;

            // 시트 탭
            _tabs = new TabControl();
            _tabs.Dock = DockStyle.Top;
            _tabs.Height = 28;
            _tabs.SelectedIndexChanged += delegate { OnSheetChanged(); };

            // 경로 라벨 패널
            Panel pathPanel = new Panel();
            pathPanel.Dock = DockStyle.Top;
            pathPanel.Height = 22;
            _lblLeftPath = new Label();
            _lblLeftPath.Text = "(좌측 파일 없음)";
            _lblLeftPath.Dock = DockStyle.Left;
            _lblLeftPath.Width = 620;
            _lblLeftPath.TextAlign = ContentAlignment.MiddleLeft;
            _lblLeftPath.AutoEllipsis = true;
            _lblRightPath = new Label();
            _lblRightPath.Text = "(우측 파일 없음)";
            _lblRightPath.Dock = DockStyle.Fill;
            _lblRightPath.TextAlign = ContentAlignment.MiddleLeft;
            _lblRightPath.AutoEllipsis = true;
            pathPanel.Controls.Add(_lblRightPath);
            pathPanel.Controls.Add(_lblLeftPath);

            // 좌우 그리드
            SplitContainer split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.SplitterWidth = 4;

            _gridLeft = MakeGrid();
            _gridRight = MakeGrid();
            split.Panel1.Controls.Add(_gridLeft);
            split.Panel2.Controls.Add(_gridRight);
            // 균등 분할.
            Load += delegate { try { split.SplitterDistance = split.Width / 2; } catch { } };

            // 상태바
            StatusStrip status = new StatusStrip();
            _lblSummary = new ToolStripStatusLabel("파일을 열고 [비교]를 누르세요.");
            _lblSummary.Spring = true;
            _lblSummary.TextAlign = ContentAlignment.MiddleLeft;
            _progress = new ToolStripProgressBar();
            _progress.Style = ProgressBarStyle.Marquee;
            _progress.Visible = false;
            status.Items.Add(_lblSummary);
            status.Items.Add(_progress);

            // 도킹 z-order: Fill(split) 을 먼저 추가하고, 안쪽→바깥쪽 순으로 Top 들을,
            // 마지막에 Bottom(status) 을 추가한다. (WinForms 도킹은 이 순서에 의존)
            Controls.Add(split);      // Fill (가장 안쪽)
            Controls.Add(pathPanel);  // Top (안쪽)
            Controls.Add(_tabs);      // Top
            Controls.Add(tool);       // Top (가장 위)
            Controls.Add(status);     // Bottom

            // 워커
            _worker = new BackgroundWorker();
            _worker.WorkerReportsProgress = true;
            _worker.DoWork += Worker_DoWork;
            _worker.RunWorkerCompleted += Worker_Completed;
        }

        private DataGridView MakeGrid()
        {
            DataGridView g = new DataGridView();
            g.Dock = DockStyle.Fill;
            g.VirtualMode = true;
            g.ReadOnly = true;
            g.AllowUserToAddRows = false;
            g.AllowUserToDeleteRows = false;
            g.AllowUserToResizeRows = false;
            g.RowHeadersWidth = 60;
            g.SelectionMode = DataGridViewSelectionMode.CellSelect;
            g.EditMode = DataGridViewEditMode.EditProgrammatically;
            g.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.None;
            g.AllowUserToResizeColumns = true;
            g.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            g.CellValueNeeded += Grid_CellValueNeeded;
            g.CellFormatting += Grid_CellFormatting;
            g.Scroll += Grid_Scroll;
            g.SelectionChanged += Grid_SelectionChanged;
            g.RowPostPaint += Grid_RowPostPaint;
            return g;
        }

        private static void AddButton(ToolStrip tool, string text, EventHandler onClick)
        {
            ToolStripButton b = new ToolStripButton(text);
            b.DisplayStyle = ToolStripItemDisplayStyle.Text;
            b.Click += onClick;
            tool.Items.Add(b);
        }

        // ---------------------------------------------------------------- 파일 열기 / 비교

        private void OpenFile(bool left)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel 파일 (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|모든 파일 (*.*)|*.*";
                string last = _settings.LastFolder;
                if (!string.IsNullOrEmpty(last) && Directory.Exists(last))
                    dlg.InitialDirectory = last;
                if (dlg.ShowDialog(this) != DialogResult.OK) return;

                if (left) { _leftPath = dlg.FileName; _lblLeftPath.Text = dlg.FileName; }
                else { _rightPath = dlg.FileName; _lblRightPath.Text = dlg.FileName; }

                _settings.LastFolder = Path.GetDirectoryName(dlg.FileName);
                _settings.Save();
            }
        }

        private void StartCompare()
        {
            if (string.IsNullOrEmpty(_leftPath) || string.IsNullOrEmpty(_rightPath))
            {
                MessageBox.Show(this, "좌측과 우측 파일을 모두 지정하세요.", "안내",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_worker.IsBusy) return;

            _pendingLeft.Clear();
            _pendingRight.Clear();
            _progress.Visible = true;
            _lblSummary.Text = "로드/비교 중…";
            SetBusy(true);
            _worker.RunWorkerAsync(new string[] { _leftPath, _rightPath });
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            string[] paths = (string[])e.Argument;
            using (ExcelComReader reader = new ExcelComReader())
            {
                WorkbookData l = reader.LoadWorkbook(paths[0]);
                WorkbookData r = reader.LoadWorkbook(paths[1]);
                e.Result = DiffEngine.Compare(l, r);
            }
        }

        private void Worker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            _progress.Visible = false;
            SetBusy(false);
            if (e.Error != null)
            {
                _lblSummary.Text = "오류: " + e.Error.Message;
                MessageBox.Show(this, DescribeError(e.Error), "비교 실패",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            _diff = (DiffResult)e.Result;
            _settings.AddRecentPair(_leftPath, _rightPath);
            _settings.Save();
            PopulateTabs();
            _lblSummary.Text = _diff.Summary();
        }

        private void SetBusy(bool busy)
        {
            if (_tool != null) _tool.Enabled = !busy;
        }

        /// <summary>DRM/COM 예외를 사용자 안내 문구로 매핑 (spec §7, §11).</summary>
        private static string DescribeError(Exception ex)
        {
            string m = ex.Message ?? "";
            if (m.IndexOf("권한", StringComparison.Ordinal) >= 0)
                return m;
            // 일반적인 COM 오류 힌트.
            return "파일을 열거나 비교하는 중 오류가 발생했습니다.\r\n\r\n"
                 + "가능 원인:\r\n"
                 + " - 파일이 이미 다른 곳에서 열려 있음\r\n"
                 + " - DRM 열람 권한이 없는 파일(우회 시도 금지)\r\n"
                 + " - Excel 이 응답하지 않음\r\n\r\n"
                 + "상세: " + m;
        }

        // ---------------------------------------------------------------- 탭 / 시트

        private void PopulateTabs()
        {
            _tabs.TabPages.Clear();
            foreach (SheetDiff sd in _diff.Sheets)
            {
                string marker = sd.HasChanges ? " ●" : "";
                TabPage tp = new TabPage(sd.Name + marker);
                tp.Tag = sd;
                _tabs.TabPages.Add(tp);
            }
            if (_tabs.TabPages.Count > 0)
            {
                _tabs.SelectedIndex = 0;
                OnSheetChanged();
            }
        }

        private void OnSheetChanged()
        {
            if (_tabs.SelectedTab == null) return;
            _curSheet = (SheetDiff)_tabs.SelectedTab.Tag;
            _curLeft = _curSheet.Left;
            _curRight = _curSheet.Right;

            // 열 범위.
            _colStart = _curSheet.MinCol;
            _colCount = Math.Max(0, _curSheet.MaxCol - _curSheet.MinCol + 1);

            // 네비 셀 정렬.
            _navCells = new List<DiffCell>(_curSheet.Cells);
            _navCells.Sort(delegate(DiffCell a, DiffCell b)
            {
                if (a.Row != b.Row) return a.Row.CompareTo(b.Row);
                return a.Col.CompareTo(b.Col);
            });
            _navIndex = -1;

            BuildColumns(_gridLeft);
            BuildColumns(_gridRight);
            RefreshGridRows();
        }

        private void BuildColumns(DataGridView g)
        {
            g.RowCount = 0;          // 가상 모드에서 열 제거 전 행 먼저 비움
            g.Columns.Clear();
            if (_colCount <= 0) return;
            DataGridViewColumn[] cols = new DataGridViewColumn[_colCount];
            for (int i = 0; i < _colCount; i++)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.HeaderText = ColLetter(_colStart + i);
                col.Width = 110;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                cols[i] = col;
            }
            g.Columns.AddRange(cols);
        }

        private void RefreshGridRows()
        {
            if (_curSheet == null) return;

            // 표시할 절대행 목록 산정.
            List<int> rows = new List<int>();
            if (_changesOnly)
            {
                HashSet<int> diffRows = new HashSet<int>();
                foreach (DiffCell dc in _curSheet.Cells) diffRows.Add(dc.Row);
                List<int> sorted = new List<int>(diffRows);
                sorted.Sort();
                rows = sorted;
            }
            else
            {
                for (int r = _curSheet.MinRow; r <= _curSheet.MaxRow; r++) rows.Add(r);
            }

            _rowMap = rows.ToArray();
            _absRowToDisplay = new Dictionary<int, int>(_rowMap.Length);
            for (int i = 0; i < _rowMap.Length; i++) _absRowToDisplay[_rowMap[i]] = i;

            SetRowCount(_gridLeft, _rowMap.Length);
            SetRowCount(_gridRight, _rowMap.Length);
            _gridLeft.Invalidate();
            _gridRight.Invalidate();
        }

        private static void SetRowCount(DataGridView g, int n)
        {
            if (g.Columns.Count == 0) { g.RowCount = 0; return; }
            g.RowCount = n;
        }

        // ---------------------------------------------------------------- 가상 그리드 콜백

        private void Grid_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (_curSheet == null) { e.Value = ""; return; }
            if (e.RowIndex < 0 || e.RowIndex >= _rowMap.Length) { e.Value = ""; return; }
            bool isLeft = (sender == _gridLeft);
            int absRow = _rowMap[e.RowIndex];
            int absCol = _colStart + e.ColumnIndex;

            object v = GetEffectiveValue(isLeft, absRow, absCol);
            e.Value = DiffEngine.ToText(v);
        }

        /// <summary>병합 대기값이 있으면 그 값을, 아니면 원본 값을 반환.</summary>
        private object GetEffectiveValue(bool isLeft, int absRow, int absCol)
        {
            string key = CellKey(_curSheet.Name, absRow, absCol);
            Dictionary<string, object> pend = isLeft ? _pendingLeft : _pendingRight;
            object pv;
            if (pend.TryGetValue(key, out pv)) return pv;

            SheetData s = isLeft ? _curLeft : _curRight;
            return s != null ? s.GetValueAbs(absRow, absCol) : null;
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (_curSheet == null) return;
            if (e.RowIndex < 0 || e.RowIndex >= _rowMap.Length) return;
            bool isLeft = (sender == _gridLeft);
            int absRow = _rowMap[e.RowIndex];
            int absCol = _colStart + e.ColumnIndex;

            // 병합 적용된 셀은 파랑으로.
            string key = CellKey(_curSheet.Name, absRow, absCol);
            if ((isLeft ? _pendingLeft : _pendingRight).ContainsKey(key))
            {
                e.CellStyle.BackColor = ColMerged;
                return;
            }

            DiffCell dc = _curSheet.GetCell(absRow, absCol);
            if (dc == null) return;

            switch (dc.Status)
            {
                case CellStatus.Changed:
                    e.CellStyle.BackColor = ColChanged;
                    break;
                case CellStatus.Added:
                    // 추가: new(우)에만 존재 → 우측 초록, 좌측 옅은 회색(빈칸 강조).
                    e.CellStyle.BackColor = isLeft ? Color.FromArgb(240, 240, 240) : ColAdded;
                    break;
                case CellStatus.Deleted:
                    e.CellStyle.BackColor = isLeft ? ColDeleted : Color.FromArgb(240, 240, 240);
                    break;
            }
        }

        // 행 머리글에 절대 행번호 표시.
        private void Grid_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            if (e.RowIndex < 0 || e.RowIndex >= _rowMap.Length) return;
            DataGridView g = (DataGridView)sender;
            string text = _rowMap[e.RowIndex].ToString();
            Rectangle rect = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y,
                g.RowHeadersWidth - 4, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, text, g.RowHeadersDefaultCellStyle.Font, rect,
                SystemColors.ControlText, TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        }

        // ---------------------------------------------------------------- 동기 스크롤 / 선택

        private void Grid_Scroll(object sender, ScrollEventArgs e)
        {
            if (_syncing) return;
            _syncing = true;
            try
            {
                DataGridView src = (DataGridView)sender;
                DataGridView dst = (src == _gridLeft) ? _gridRight : _gridLeft;
                if (dst.RowCount == 0) return;

                if (e.ScrollOrientation == ScrollOrientation.VerticalScroll)
                {
                    int first = src.FirstDisplayedScrollingRowIndex;
                    if (first >= 0 && first < dst.RowCount)
                        dst.FirstDisplayedScrollingRowIndex = first;
                }
                else
                {
                    dst.HorizontalScrollingOffset = e.NewValue;
                }
            }
            catch { }
            finally { _syncing = false; }
        }

        private void Grid_SelectionChanged(object sender, EventArgs e)
        {
            if (_syncing) return;
            _syncing = true;
            try
            {
                DataGridView src = (DataGridView)sender;
                DataGridView dst = (src == _gridLeft) ? _gridRight : _gridLeft;
                if (src.CurrentCell == null) return;
                int r = src.CurrentCell.RowIndex;
                int c = src.CurrentCell.ColumnIndex;
                if (r >= 0 && r < dst.RowCount && c >= 0 && c < dst.ColumnCount)
                    dst.CurrentCell = dst.Rows[r].Cells[c];
            }
            catch { }
            finally { _syncing = false; }
        }

        // ---------------------------------------------------------------- 네비게이션

        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F7) { NavigateDiff(1); e.Handled = true; }
            else if (e.KeyCode == Keys.F8) { NavigateDiff(-1); e.Handled = true; }
        }

        private void NavigateDiff(int dir)
        {
            if (_curSheet == null || _navCells.Count == 0) return;
            _navIndex += dir;
            if (_navIndex < 0) _navIndex = _navCells.Count - 1;
            if (_navIndex >= _navCells.Count) _navIndex = 0;

            DiffCell dc = _navCells[_navIndex];
            int displayRow;
            if (!_absRowToDisplay.TryGetValue(dc.Row, out displayRow)) return;
            int displayCol = dc.Col - _colStart;
            if (displayCol < 0 || displayCol >= _colCount) return;

            try
            {
                _gridLeft.CurrentCell = _gridLeft.Rows[displayRow].Cells[displayCol];
                _gridLeft.FirstDisplayedScrollingRowIndex = Math.Max(0, displayRow - 3);
            }
            catch { }
            _lblSummary.Text = string.Format("{0}  |  차이 {1}/{2}  (행 {3}, 열 {4})",
                _diff.Summary(), _navIndex + 1, _navCells.Count, dc.Row, ColLetter(dc.Col));
        }

        // ---------------------------------------------------------------- 병합

        /// <summary>선택 셀에 대해 방향 병합을 대기목록에 등록. leftToRight=true → 좌 값을 우로.</summary>
        private void MergeSelected(bool leftToRight)
        {
            if (_curSheet == null) return;
            DataGridView g = leftToRight ? _gridLeft : _gridRight; // 소스 그리드
            if (g.SelectedCells.Count == 0 && g.CurrentCell != null)
                g.CurrentCell.Selected = true;

            int applied = 0;
            foreach (DataGridViewCell cell in g.SelectedCells)
            {
                if (cell.RowIndex < 0 || cell.RowIndex >= _rowMap.Length) continue;
                int absRow = _rowMap[cell.RowIndex];
                int absCol = _colStart + cell.ColumnIndex;

                object srcVal = GetEffectiveValue(leftToRight, absRow, absCol);
                string key = CellKey(_curSheet.Name, absRow, absCol);
                if (leftToRight) _pendingRight[key] = srcVal;
                else _pendingLeft[key] = srcVal;
                applied++;
            }
            _gridLeft.Invalidate();
            _gridRight.Invalidate();
            _lblSummary.Text = string.Format("{0}  |  병합 대기 좌:{1} 우:{2} (이번 {3}개 {4})",
                _diff.Summary(), _pendingLeft.Count, _pendingRight.Count, applied,
                leftToRight ? "좌→우" : "우→좌");
        }

        private void SaveMerges()
        {
            if (_pendingLeft.Count == 0 && _pendingRight.Count == 0)
            {
                MessageBox.Show(this, "저장할 병합 변경이 없습니다.", "안내",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            bool backup = MessageBox.Show(this,
                "원본 저장 전 백업 복사본을 만들까요?\r\n(권장: 예)",
                "백업", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;

            try
            {
                SetBusy(true);
                if (_pendingRight.Count > 0)
                    SaveOneSide(_rightPath, _pendingRight, backup);
                if (_pendingLeft.Count > 0)
                    SaveOneSide(_leftPath, _pendingLeft, backup);

                _pendingLeft.Clear();
                _pendingRight.Clear();
                _lblSummary.Text = "저장 완료. 다시 [비교]로 재검증을 권장합니다.";
                MessageBox.Show(this, "저장이 완료되었습니다.", "완료",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "저장 실패(원본은 보존됨):\r\n" + ex.Message +
                    "\r\n\r\nDRM 편집권한이 없거나 파일이 잠겨있을 수 있습니다.",
                    "저장 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
                _gridLeft.Invalidate();
                _gridRight.Invalidate();
            }
        }

        private void SaveOneSide(string path, Dictionary<string, object> pending, bool backup)
        {
            List<MergeItem> items = new List<MergeItem>();
            foreach (KeyValuePair<string, object> kv in pending)
            {
                // key = "sheet:row:col". 시트명에는 ':' 가 올 수 없으므로 뒤 두 조각이 row/col.
                int lastColon = kv.Key.LastIndexOf(':');
                int prevColon = kv.Key.LastIndexOf(':', lastColon - 1);
                string sheet = kv.Key.Substring(0, prevColon);
                int row = int.Parse(kv.Key.Substring(prevColon + 1, lastColon - prevColon - 1));
                int col = int.Parse(kv.Key.Substring(lastColon + 1));
                items.Add(new MergeItem(sheet, row, col, kv.Value));
            }
            using (MergeEngine me = new MergeEngine())
            {
                string backupPath;
                me.ApplyAndSave(path, items, backup, out backupPath);
            }
        }

        // ---------------------------------------------------------------- N-way

        private void OpenNWayDialog()
        {
            using (NWayForm f = new NWayForm(_settings))
            {
                f.ShowDialog(this);
            }
        }

        // ---------------------------------------------------------------- 드래그앤드롭

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
                e.Effect = DragDropEffects.Copy;
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0) return;
            if (files.Length >= 2)
            {
                _leftPath = files[0]; _lblLeftPath.Text = files[0];
                _rightPath = files[1]; _lblRightPath.Text = files[1];
                StartCompare();
            }
            else
            {
                if (string.IsNullOrEmpty(_leftPath)) { _leftPath = files[0]; _lblLeftPath.Text = files[0]; }
                else { _rightPath = files[0]; _lblRightPath.Text = files[0]; }
            }
        }

        // ---------------------------------------------------------------- 온보딩

        private void ShowOnboardingIfNeeded()
        {
            if (_settings.HideOnboarding) return;
            Shown += delegate
            {
                using (Form f = new Form())
                {
                    f.Text = "환영합니다 — Excel Diff / Merge";
                    f.Width = 560; f.Height = 320;
                    f.StartPosition = FormStartPosition.CenterParent;
                    f.FormBorderStyle = FormBorderStyle.FixedDialog;
                    f.MaximizeBox = false; f.MinimizeBox = false;

                    Label lbl = new Label();
                    lbl.Dock = DockStyle.Fill;
                    lbl.Padding = new Padding(16);
                    lbl.Text =
                        "사용법\r\n\r\n"
                      + "1) [좌측 열기] / [우측 열기] 로 비교할 엑셀 파일 2개 선택\r\n"
                      + "   (또는 파일 2개를 창에 드래그앤드롭)\r\n"
                      + "2) [비교] 클릭 → 좌우 병렬로 차이 표시\r\n"
                      + "     노랑=변경  초록=추가  빨강=삭제\r\n"
                      + "3) F7/F8 로 다음/이전 차이 이동\r\n"
                      + "4) 셀 선택 후 [좌→우]/[우→좌] 로 병합, [저장]\r\n\r\n"
                      + "* DRM 파일은 Excel 이 복호화하여 읽습니다(우회 없음).\r\n"
                      + "* 편집권한 없는 파일은 저장이 거부될 수 있습니다.";
                    f.Controls.Add(lbl);

                    CheckBox chk = new CheckBox();
                    chk.Text = "다시 보지 않기";
                    chk.Dock = DockStyle.Bottom;
                    chk.Height = 30;
                    chk.Padding = new Padding(16, 0, 0, 0);
                    f.Controls.Add(chk);

                    Button ok = new Button();
                    ok.Text = "시작";
                    ok.Dock = DockStyle.Bottom;
                    ok.Height = 34;
                    ok.Click += delegate { f.Close(); };
                    f.Controls.Add(ok);

                    f.ShowDialog(this);
                    if (chk.Checked) { _settings.HideOnboarding = true; _settings.Save(); }
                }
            };
        }

        // ---------------------------------------------------------------- 유틸

        private static string CellKey(string sheet, int row, int col)
        {
            return sheet + ":" + row + ":" + col;
        }

        /// <summary>열 번호(1-based) → Excel 열문자(A, B, …, AA).</summary>
        public static string ColLetter(int col)
        {
            string s = "";
            while (col > 0)
            {
                int rem = (col - 1) % 26;
                s = (char)('A' + rem) + s;
                col = (col - 1) / 26;
            }
            return s.Length == 0 ? "?" : s;
        }
    }
}
