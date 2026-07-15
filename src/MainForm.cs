using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelDiffMerge
{
    /// <summary>
    /// WinMerge 식 좌우 병렬 뷰 (spec §6 + 시장조사 UX 반영).
    ///  - DataGridView 2개(가상모드) + 색상 + 동기 스크롤 + 시트 탭.
    ///  - 행 정렬(좌표/자동LCS/키컬럼) → 삽입/삭제 cascade 방지.
    ///  - 변경 네비(F7/F8), 요약, "변경만 보기" 필터, 셀 상세 패널, 색상 범례, 변경 마커바.
    ///  - 방향 병합(좌→우/우→좌) → 저장(백업). CSV 폴더 비교(로컬 테스트).
    ///  - COM 로드는 STA 전용 스레드에서 수행(조사 반영), UI 는 Invoke 로 갱신.
    /// </summary>
    public sealed class MainForm : Form
    {
        // 색상
        private static readonly Color ColChanged = Color.FromArgb(255, 243, 160); // 노랑
        private static readonly Color ColAdded = Color.FromArgb(186, 238, 186);   // 초록
        private static readonly Color ColDeleted = Color.FromArgb(250, 186, 186); // 빨강
        private static readonly Color ColMerged = Color.FromArgb(186, 212, 255);  // 병합됨(파랑)
        private static readonly Color ColGap = Color.FromArgb(238, 238, 238);     // 반대편 없는 셀

        private readonly Settings _settings;

        private ToolStrip _tool;
        private ToolStripComboBox _cboAlign;
        private ToolStripTextBox _txtKeyCol;
        private DataGridView _gridLeft;
        private DataGridView _gridRight;
        private MarkerBar _marker;
        private TabControl _tabs;
        private Label _lblLeftPath;
        private Label _lblRightPath;
        private ToolStripStatusLabel _lblSummary;
        private ToolStripProgressBar _progress;
        private ToolStripButton _btnChangesOnly;
        private Label _detail;

        // 상태
        private string _leftPath;
        private string _rightPath;
        private WorkbookData _leftWb;   // 로드 캐시(정렬 변경 시 재비교만)
        private WorkbookData _rightWb;
        private DiffResult _diff;
        private SheetDiff _curSheet;

        private int[] _rowMap = new int[0];             // 표시행 → _curSheet.Rows 인덱스
        private Dictionary<int, int> _rowsIdxToDisplay = new Dictionary<int, int>();
        private int _colStart = 1;
        private int _colCount;

        private AlignMode _alignMode = AlignMode.Auto;
        private int _keyCol = -1;
        private bool _changesOnly;
        private bool _syncing;
        private bool _busy;

        private readonly Dictionary<string, object> _pendingLeft = new Dictionary<string, object>();
        private readonly Dictionary<string, object> _pendingRight = new Dictionary<string, object>();

        private int _navIndex = -1;

        public MainForm()
        {
            _settings = Settings.LoadDefault();
            _alignMode = (AlignMode)IntSetting("AlignMode", (int)AlignMode.Auto);
            BuildUi();
            ShowOnboardingIfNeeded();
        }

        private int IntSetting(string key, int def)
        {
            int v;
            return int.TryParse(_settings.Get(key, ""), out v) ? v : def;
        }

        // ================================================================ UI
        private void BuildUi()
        {
            Text = "Excel Diff / Merge";
            Width = 1320; Height = 860;
            StartPosition = FormStartPosition.CenterScreen;
            AllowDrop = true;
            DragEnter += OnDragEnter;
            DragDrop += OnDragDrop;
            KeyPreview = true;
            KeyDown += OnKeyDown;

            // ----- 툴바
            ToolStrip tool = new ToolStrip();
            _tool = tool;
            tool.GripStyle = ToolStripGripStyle.Hidden;
            AddButton(tool, "좌측 열기", delegate { OpenFile(true); });
            AddButton(tool, "우측 열기", delegate { OpenFile(false); });
            AddButton(tool, "비교", delegate { StartCompareExcel(); });
            tool.Items.Add(new ToolStripSeparator());
            tool.Items.Add(new ToolStripLabel("정렬:"));
            _cboAlign = new ToolStripComboBox();
            _cboAlign.DropDownStyle = ComboBoxStyle.DropDownList;
            _cboAlign.Items.AddRange(new object[] { "좌표(빠름)", "자동정렬(LCS)", "키 컬럼" });
            _cboAlign.SelectedIndex = (int)_alignMode;
            _cboAlign.SelectedIndexChanged += delegate { OnAlignModeChanged(); };
            tool.Items.Add(_cboAlign);
            tool.Items.Add(new ToolStripLabel("키열:"));
            _txtKeyCol = new ToolStripTextBox();
            _txtKeyCol.Width = 40;
            _txtKeyCol.ToolTipText = "키 컬럼 문자(예: A). '키 컬럼' 정렬에서 사용.";
            _txtKeyCol.LostFocus += delegate { OnKeyColChanged(); };
            tool.Items.Add(_txtKeyCol);
            tool.Items.Add(new ToolStripSeparator());
            AddButton(tool, "◀이전(F8)", delegate { NavigateDiff(-1); });
            AddButton(tool, "다음(F7)▶", delegate { NavigateDiff(1); });
            _btnChangesOnly = new ToolStripButton("변경만");
            _btnChangesOnly.CheckOnClick = true;
            _btnChangesOnly.CheckedChanged += delegate { _changesOnly = _btnChangesOnly.Checked; RefreshGridRows(); };
            tool.Items.Add(_btnChangesOnly);
            tool.Items.Add(new ToolStripSeparator());
            AddButton(tool, "좌→우", delegate { MergeSelected(true); });
            AddButton(tool, "우→좌", delegate { MergeSelected(false); });
            AddButton(tool, "저장", delegate { SaveMerges(); });
            tool.Items.Add(new ToolStripSeparator());
            AddButton(tool, "N-way…", delegate { OpenNWayDialog(); });
            AddButton(tool, "CSV폴더비교", delegate { StartCompareCsv(); });
            tool.Dock = DockStyle.Top;

            // ----- 범례
            Panel legend = BuildLegend();

            // ----- 시트 탭
            _tabs = new TabControl();
            _tabs.Dock = DockStyle.Top;
            _tabs.Height = 26;
            _tabs.SelectedIndexChanged += delegate { OnSheetChanged(); };

            // ----- 경로 라벨
            Panel pathPanel = new Panel();
            pathPanel.Dock = DockStyle.Top;
            pathPanel.Height = 22;
            _lblLeftPath = new Label();
            _lblLeftPath.Text = "(좌측 파일 없음)";
            _lblLeftPath.Dock = DockStyle.Left; _lblLeftPath.Width = 640;
            _lblLeftPath.TextAlign = ContentAlignment.MiddleLeft; _lblLeftPath.AutoEllipsis = true;
            _lblRightPath = new Label();
            _lblRightPath.Text = "(우측 파일 없음)";
            _lblRightPath.Dock = DockStyle.Fill;
            _lblRightPath.TextAlign = ContentAlignment.MiddleLeft; _lblRightPath.AutoEllipsis = true;
            pathPanel.Controls.Add(_lblRightPath);
            pathPanel.Controls.Add(_lblLeftPath);

            // ----- 좌우 그리드 + 마커바
            SplitContainer split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.SplitterWidth = 4;
            _gridLeft = MakeGrid();
            _gridRight = MakeGrid();
            split.Panel1.Controls.Add(_gridLeft);

            _marker = new MarkerBar();
            _marker.Dock = DockStyle.Right;
            _marker.OnSeek = delegate(float pos) { SeekTo(pos); };
            split.Panel2.Controls.Add(_gridRight);
            split.Panel2.Controls.Add(_marker);

            Load += delegate { try { split.SplitterDistance = split.Width / 2; } catch { } };

            // ----- 셀 상세 패널
            Panel detailPanel = new Panel();
            detailPanel.Dock = DockStyle.Bottom;
            detailPanel.Height = 84;
            _detail = new Label();
            _detail.Dock = DockStyle.Fill;
            _detail.Font = new Font(FontFamily.GenericMonospace, 9f);
            _detail.Padding = new Padding(8, 4, 8, 4);
            _detail.Text = "셀을 선택하면 좌/우 값·수식을 여기에 표시합니다.";
            detailPanel.Controls.Add(_detail);

            // ----- 상태바
            StatusStrip status = new StatusStrip();
            _lblSummary = new ToolStripStatusLabel("파일을 열고 [비교]를 누르세요.");
            _lblSummary.Spring = true; _lblSummary.TextAlign = ContentAlignment.MiddleLeft;
            _progress = new ToolStripProgressBar();
            _progress.Style = ProgressBarStyle.Marquee; _progress.Visible = false;
            status.Items.Add(_lblSummary);
            status.Items.Add(_progress);

            // 도킹 z-order: Fill 먼저, 안쪽→바깥쪽 Top, 마지막에 Bottom.
            Controls.Add(split);        // Fill
            Controls.Add(pathPanel);    // Top
            Controls.Add(_tabs);        // Top
            Controls.Add(legend);       // Top
            Controls.Add(tool);         // Top(최상단)
            Controls.Add(detailPanel);  // Bottom
            Controls.Add(status);       // Bottom(최하단)

            UpdateKeyColEnabled();
        }

        private Panel BuildLegend()
        {
            Panel legend = new Panel();
            legend.Dock = DockStyle.Top;
            legend.Height = 22;
            int x = 8;
            x = AddLegendItem(legend, x, ColChanged, "변경");
            x = AddLegendItem(legend, x, ColAdded, "추가");
            x = AddLegendItem(legend, x, ColDeleted, "삭제");
            x = AddLegendItem(legend, x, ColMerged, "병합됨");
            x = AddLegendItem(legend, x, ColGap, "빈칸(정렬)");
            return legend;
        }

        private static int AddLegendItem(Panel host, int x, Color color, string text)
        {
            Panel sw = new Panel();
            sw.BackColor = color;
            sw.BorderStyle = BorderStyle.FixedSingle;
            sw.SetBounds(x, 4, 16, 14);
            host.Controls.Add(sw);
            Label lb = new Label();
            lb.Text = text; lb.AutoSize = true;
            lb.SetBounds(x + 20, 4, 10, 14);
            host.Controls.Add(lb);
            return x + 24 + text.Length * 12 + 12;
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
            g.RowHeadersWidth = 64;
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

        // ================================================================ 파일 열기/비교
        private void OpenFile(bool left)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel 파일 (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls|모든 파일 (*.*)|*.*";
                string last = _settings.LastFolder;
                if (!string.IsNullOrEmpty(last) && Directory.Exists(last)) dlg.InitialDirectory = last;
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                if (left) { _leftPath = dlg.FileName; _lblLeftPath.Text = dlg.FileName; }
                else { _rightPath = dlg.FileName; _lblRightPath.Text = dlg.FileName; }
                _settings.LastFolder = Path.GetDirectoryName(dlg.FileName);
                _settings.Save();
            }
        }

        private void StartCompareExcel()
        {
            if (string.IsNullOrEmpty(_leftPath) || string.IsNullOrEmpty(_rightPath))
            {
                Info("좌측과 우측 파일을 모두 지정하세요."); return;
            }
            LoadAndCompare(false);
        }

        private void StartCompareCsv()
        {
            string l = PickFolder("좌측 CSV 폴더(각 .csv = 시트)");
            if (l == null) return;
            string r = PickFolder("우측 CSV 폴더");
            if (r == null) return;
            _leftPath = l; _rightPath = r;
            _lblLeftPath.Text = l; _lblRightPath.Text = r;
            LoadAndCompare(true);
        }

        private string PickFolder(string title)
        {
            using (FolderBrowserDialog dlg = new FolderBrowserDialog())
            {
                dlg.Description = title;
                if (dlg.ShowDialog(this) != DialogResult.OK) return null;
                return dlg.SelectedPath;
            }
        }

        /// <summary>로드(STA 스레드) → 캐시 → RunDiff. useCsv=true 면 CSV 리더.</summary>
        private void LoadAndCompare(bool useCsv)
        {
            if (_busy) return;
            _pendingLeft.Clear();
            _pendingRight.Clear();
            _progress.Visible = true;
            _lblSummary.Text = "로드 중…";
            SetBusy(true);

            string lp = _leftPath, rp = _rightPath;
            StaTask.Run<WorkbookData[]>(
                delegate
                {
                    using (IWorkbookReader reader = useCsv ? (IWorkbookReader)new CsvWorkbookReader() : new ExcelComReader())
                    {
                        WorkbookData l = reader.LoadWorkbook(lp);
                        WorkbookData r = reader.LoadWorkbook(rp);
                        return new WorkbookData[] { l, r };
                    }
                },
                delegate(WorkbookData[] res, Exception err)
                {
                    // STA 워커 스레드 → UI 스레드로 마샬링.
                    BeginInvoke((MethodInvoker)delegate { OnLoaded(res, err); });
                });
        }

        private void OnLoaded(WorkbookData[] res, Exception err)
        {
            _progress.Visible = false;
            SetBusy(false);
            if (err != null)
            {
                _lblSummary.Text = "오류: " + err.Message;
                MessageBox.Show(this, DescribeError(err), "로드 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            _leftWb = res[0];
            _rightWb = res[1];
            _settings.AddRecentPair(_leftPath, _rightPath);
            _settings.Save();
            RunDiff();
        }

        /// <summary>캐시된 워크북으로 정렬/키 반영해 diff 재계산(재로드 없음).</summary>
        private void RunDiff()
        {
            if (_leftWb == null || _rightWb == null) return;
            _pendingLeft.Clear();
            _pendingRight.Clear();
            try
            {
                _diff = DiffEngine.Compare(_leftWb, _rightWb, _alignMode, _keyCol);
            }
            catch (Exception ex)
            {
                _lblSummary.Text = "비교 오류: " + ex.Message;
                return;
            }
            PopulateTabs();
            _lblSummary.Text = _diff.Summary() + "  [" + AlignName() + "]";
        }

        private string AlignName()
        {
            if (_alignMode == AlignMode.Coordinate) return "좌표";
            if (_alignMode == AlignMode.KeyColumn) return "키:" + (_keyCol >= 0 ? ColLetter(_keyCol) : "?");
            return "자동정렬";
        }

        private void OnAlignModeChanged()
        {
            _alignMode = (AlignMode)_cboAlign.SelectedIndex;
            _settings.Set("AlignMode", ((int)_alignMode).ToString());
            _settings.Save();
            UpdateKeyColEnabled();
            RunDiff();
        }

        private void OnKeyColChanged()
        {
            _keyCol = ParseColLetter(_txtKeyCol.Text);
            if (_alignMode == AlignMode.KeyColumn) RunDiff();
        }

        private void UpdateKeyColEnabled()
        {
            bool key = _alignMode == AlignMode.KeyColumn;
            _txtKeyCol.Enabled = key;
        }

        private void SetBusy(bool busy)
        {
            _busy = busy;
            if (_tool != null) _tool.Enabled = !busy;
        }

        private static string DescribeError(Exception ex)
        {
            string m = ex.Message ?? "";
            if (m.IndexOf("권한", StringComparison.Ordinal) >= 0) return m;
            return "파일을 열거나 비교하는 중 오류가 발생했습니다.\r\n\r\n"
                 + "가능 원인:\r\n"
                 + " - 파일이 이미 다른 곳에서 열려 있음\r\n"
                 + " - DRM 열람 권한이 없는 파일(우회 시도 금지)\r\n"
                 + " - Excel 이 응답하지 않음\r\n\r\n상세: " + m;
        }

        // ================================================================ 탭/시트
        private void PopulateTabs()
        {
            _tabs.TabPages.Clear();
            foreach (SheetDiff sd in _diff.Sheets)
            {
                string suffix = "";
                if (sd.Status == SheetStatus.Added) suffix = " ＋";
                else if (sd.Status == SheetStatus.Deleted) suffix = " －";
                else if (sd.HasChanges) suffix = " ●(" + (sd.ChangedCount + sd.AddedCount + sd.DeletedCount) + ")";
                TabPage tp = new TabPage(sd.Name + suffix);
                tp.Tag = sd;
                _tabs.TabPages.Add(tp);
            }
            if (_tabs.TabPages.Count > 0)
            {
                // 첫 변경 시트로 이동(없으면 0).
                int idx = 0;
                for (int i = 0; i < _diff.Sheets.Count; i++)
                    if (_diff.Sheets[i].HasChanges) { idx = i; break; }
                _tabs.SelectedIndex = idx;
                OnSheetChanged();
            }
        }

        private void OnSheetChanged()
        {
            if (_tabs.SelectedTab == null) return;
            _curSheet = (SheetDiff)_tabs.SelectedTab.Tag;
            _colStart = _curSheet.MinCol;
            _colCount = _curSheet.ColCount;
            _navIndex = -1;

            BuildColumns(_gridLeft);
            BuildColumns(_gridRight);
            RefreshGridRows();
        }

        private void BuildColumns(DataGridView g)
        {
            g.RowCount = 0;
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
            List<int> rows = new List<int>();
            for (int i = 0; i < _curSheet.Rows.Count; i++)
            {
                if (_changesOnly && _curSheet.Rows[i].Kind == RowKind.Same) continue;
                rows.Add(i);
            }
            _rowMap = rows.ToArray();
            _rowsIdxToDisplay = new Dictionary<int, int>(_rowMap.Length);
            for (int i = 0; i < _rowMap.Length; i++) _rowsIdxToDisplay[_rowMap[i]] = i;

            SetRowCount(_gridLeft, _rowMap.Length);
            SetRowCount(_gridRight, _rowMap.Length);
            _gridLeft.Invalidate();
            _gridRight.Invalidate();
            UpdateMarkerBar();
        }

        private static void SetRowCount(DataGridView g, int n)
        {
            if (g.Columns.Count == 0) { g.RowCount = 0; return; }
            g.RowCount = n;
        }

        // ================================================================ 가상 그리드
        private void Grid_CellValueNeeded(object sender, DataGridViewCellValueEventArgs e)
        {
            if (_curSheet == null || e.RowIndex < 0 || e.RowIndex >= _rowMap.Length) { e.Value = ""; return; }
            bool isLeft = (sender == _gridLeft);
            DiffRow dr = _curSheet.Rows[_rowMap[e.RowIndex]];
            int absCol = _colStart + e.ColumnIndex;
            e.Value = DiffEngine.ToText(EffectiveValue(isLeft, dr, absCol));
        }

        private object EffectiveValue(bool isLeft, DiffRow dr, int absCol)
        {
            int absRow = isLeft ? dr.LeftRow : dr.RightRow;
            if (absRow < 0) return null;
            string key = CellKey(_curSheet.Name, absRow, absCol);
            Dictionary<string, object> pend = isLeft ? _pendingLeft : _pendingRight;
            object pv;
            if (pend.TryGetValue(key, out pv)) return pv;
            SheetData s = isLeft ? _curSheet.Left : _curSheet.Right;
            return s != null ? s.GetValueAbs(absRow, absCol) : null;
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (_curSheet == null || e.RowIndex < 0 || e.RowIndex >= _rowMap.Length) return;
            bool isLeft = (sender == _gridLeft);
            DiffRow dr = _curSheet.Rows[_rowMap[e.RowIndex]];
            int absCol = _colStart + e.ColumnIndex;

            int absRow = isLeft ? dr.LeftRow : dr.RightRow;
            if (absRow >= 0)
            {
                string key = CellKey(_curSheet.Name, absRow, absCol);
                if ((isLeft ? _pendingLeft : _pendingRight).ContainsKey(key))
                {
                    e.CellStyle.BackColor = ColMerged; return;
                }
            }
            else
            {
                // 이 행에 이쪽 셀이 없음(반대편만) → 빈칸 표시.
                e.CellStyle.BackColor = ColGap; return;
            }

            CellStatus st = dr.StatusAt(absCol);
            switch (st)
            {
                case CellStatus.Changed: e.CellStyle.BackColor = ColChanged; break;
                case CellStatus.Added: e.CellStyle.BackColor = isLeft ? ColGap : ColAdded; break;
                case CellStatus.Deleted: e.CellStyle.BackColor = isLeft ? ColDeleted : ColGap; break;
            }
        }

        private void Grid_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            if (_curSheet == null || e.RowIndex < 0 || e.RowIndex >= _rowMap.Length) return;
            DataGridView g = (DataGridView)sender;
            DiffRow dr = _curSheet.Rows[_rowMap[e.RowIndex]];
            int absRow = (g == _gridLeft) ? dr.LeftRow : dr.RightRow;
            string text = absRow >= 0 ? absRow.ToString() : "·";
            Rectangle rect = new Rectangle(e.RowBounds.Location.X, e.RowBounds.Location.Y,
                g.RowHeadersWidth - 4, e.RowBounds.Height);
            TextRenderer.DrawText(e.Graphics, text, g.RowHeadersDefaultCellStyle.Font, rect,
                SystemColors.ControlText, TextFormatFlags.Right | TextFormatFlags.VerticalCenter);
        }

        // ================================================================ 동기 스크롤/선택
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
                    if (first >= 0 && first < dst.RowCount) dst.FirstDisplayedScrollingRowIndex = first;
                }
                else
                {
                    dst.HorizontalScrollingOffset = e.NewValue;
                }
            }
            catch { }
            finally { _syncing = false; UpdateMarkerViewport(); }
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
                int r = src.CurrentCell.RowIndex, c = src.CurrentCell.ColumnIndex;
                if (r >= 0 && r < dst.RowCount && c >= 0 && c < dst.ColumnCount)
                    dst.CurrentCell = dst.Rows[r].Cells[c];
                UpdateDetail(r, c);
            }
            catch { }
            finally { _syncing = false; }
        }

        private void UpdateDetail(int displayRow, int col)
        {
            if (_curSheet == null || displayRow < 0 || displayRow >= _rowMap.Length || col < 0)
            {
                _detail.Text = ""; return;
            }
            DiffRow dr = _curSheet.Rows[_rowMap[displayRow]];
            int absCol = _colStart + col;
            object lv = dr.LeftRow >= 0 && _curSheet.Left != null ? _curSheet.Left.GetValueAbs(dr.LeftRow, absCol) : null;
            object rv = dr.RightRow >= 0 && _curSheet.Right != null ? _curSheet.Right.GetValueAbs(dr.RightRow, absCol) : null;
            object lf = dr.LeftRow >= 0 && _curSheet.Left != null ? _curSheet.Left.GetFormulaAbs(dr.LeftRow, absCol) : null;
            object rf = dr.RightRow >= 0 && _curSheet.Right != null ? _curSheet.Right.GetFormulaAbs(dr.RightRow, absCol) : null;
            CellStatus st = dr.StatusAt(absCol);

            string addr = ColLetter(absCol);
            string lrow = dr.LeftRow >= 0 ? dr.LeftRow.ToString() : "-";
            string rrow = dr.RightRow >= 0 ? dr.RightRow.ToString() : "-";
            _detail.Text =
                string.Format("[{0}]  {1}열  (좌 {2}행 / 우 {3}행)\r\n", KindLabel(st), addr, lrow, rrow)
              + string.Format("좌 값: {0}   |   좌 수식: {1}\r\n", Trunc(DiffEngine.ToText(lv)), Trunc(FormulaText(lf)))
              + string.Format("우 값: {0}   |   우 수식: {1}", Trunc(DiffEngine.ToText(rv)), Trunc(FormulaText(rf)));
        }

        private static string FormulaText(object f)
        {
            string s = f == null ? "" : f.ToString();
            return (s.Length > 0 && s[0] == '=') ? s : "(수식 아님)";
        }

        private static string Trunc(string s)
        {
            if (s == null) return "";
            return s.Length > 120 ? s.Substring(0, 120) + "…" : s;
        }

        private static string KindLabel(CellStatus st)
        {
            switch (st)
            {
                case CellStatus.Changed: return "변경";
                case CellStatus.Added: return "추가";
                case CellStatus.Deleted: return "삭제";
                default: return "동일";
            }
        }

        // ================================================================ 마커바
        private void UpdateMarkerBar()
        {
            if (_marker == null || _curSheet == null || _rowMap.Length == 0)
            {
                if (_marker != null) _marker.SetMarks(null);
                return;
            }
            List<KeyValuePair<float, Color>> marks = new List<KeyValuePair<float, Color>>();
            int total = _rowMap.Length;
            for (int i = 0; i < total; i++)
            {
                DiffRow dr = _curSheet.Rows[_rowMap[i]];
                if (dr.Kind == RowKind.Same) continue;
                Color c = dr.Kind == RowKind.Added ? ColAdded
                        : dr.Kind == RowKind.Deleted ? ColDeleted : ColChanged;
                marks.Add(new KeyValuePair<float, Color>((float)i / Math.Max(1, total - 1), c));
            }
            _marker.SetMarks(marks);
            UpdateMarkerViewport();
        }

        private void UpdateMarkerViewport()
        {
            if (_marker == null || _gridRight.RowCount == 0) return;
            int first = _gridRight.FirstDisplayedScrollingRowIndex;
            if (first < 0) return;
            _marker.SetViewport((float)first / Math.Max(1, _gridRight.RowCount - 1));
        }

        private void SeekTo(float pos)
        {
            if (_gridRight.RowCount == 0) return;
            int row = (int)(pos * (_gridRight.RowCount - 1));
            try
            {
                _gridRight.FirstDisplayedScrollingRowIndex = Math.Max(0, row);
                _gridLeft.FirstDisplayedScrollingRowIndex = Math.Max(0, row);
            }
            catch { }
            UpdateMarkerViewport();
        }

        // ================================================================ 네비게이션
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F7) { NavigateDiff(1); e.Handled = true; }
            else if (e.KeyCode == Keys.F8) { NavigateDiff(-1); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.S) { SaveMerges(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.O) { OpenFile(true); e.Handled = true; }
        }

        private void NavigateDiff(int dir)
        {
            if (_curSheet == null || _curSheet.Nav.Count == 0) return;
            _navIndex += dir;
            if (_navIndex < 0) _navIndex = _curSheet.Nav.Count - 1;
            if (_navIndex >= _curSheet.Nav.Count) _navIndex = 0;

            DiffNav nv = _curSheet.Nav[_navIndex];
            int displayRow;
            if (!_rowsIdxToDisplay.TryGetValue(nv.RowIndex, out displayRow))
            {
                // "변경만 보기" 꺼짐 상태에서만 매핑 실패 가능 — 필터 끄고 재시도.
                if (_changesOnly)
                {
                    _btnChangesOnly.Checked = false;
                    _changesOnly = false;
                    RefreshGridRows();
                    if (!_rowsIdxToDisplay.TryGetValue(nv.RowIndex, out displayRow)) return;
                }
                else return;
            }
            int displayCol = nv.Col - _colStart;
            if (displayCol < 0 || displayCol >= _colCount) return;
            try
            {
                _gridLeft.CurrentCell = _gridLeft.Rows[displayRow].Cells[displayCol];
                _gridLeft.FirstDisplayedScrollingRowIndex = Math.Max(0, displayRow - 3);
            }
            catch { }
            _lblSummary.Text = string.Format("{0}  |  차이 {1}/{2}  ({3}열, 좌{4}/우{5}행)",
                _diff.Summary(), _navIndex + 1, _curSheet.Nav.Count, ColLetter(nv.Col),
                RowLabel(true, nv.RowIndex), RowLabel(false, nv.RowIndex));
        }

        private string RowLabel(bool left, int rowsIdx)
        {
            DiffRow dr = _curSheet.Rows[rowsIdx];
            int r = left ? dr.LeftRow : dr.RightRow;
            return r >= 0 ? r.ToString() : "-";
        }

        // ================================================================ 병합
        private void MergeSelected(bool leftToRight)
        {
            if (_curSheet == null) return;
            DataGridView g = leftToRight ? _gridLeft : _gridRight; // 소스
            if (g.SelectedCells.Count == 0 && g.CurrentCell != null) g.CurrentCell.Selected = true;

            int applied = 0, skipped = 0;
            foreach (DataGridViewCell cell in g.SelectedCells)
            {
                if (cell.RowIndex < 0 || cell.RowIndex >= _rowMap.Length) continue;
                DiffRow dr = _curSheet.Rows[_rowMap[cell.RowIndex]];
                int absCol = _colStart + cell.ColumnIndex;

                int srcRow = leftToRight ? dr.LeftRow : dr.RightRow;
                int dstRow = leftToRight ? dr.RightRow : dr.LeftRow;
                if (dstRow < 0)
                {
                    // 대상 행이 없음(행 추가/삭제 병합) → 1차 미지원.
                    skipped++;
                    continue;
                }
                object srcVal = srcRow >= 0 ? EffectiveValue(leftToRight, dr, absCol) : null;
                string key = CellKey(_curSheet.Name, dstRow, absCol);
                if (leftToRight) _pendingRight[key] = srcVal;
                else _pendingLeft[key] = srcVal;
                applied++;
            }
            _gridLeft.Invalidate();
            _gridRight.Invalidate();
            string extra = skipped > 0 ? string.Format("  (행추가/삭제 {0}건은 미지원)", skipped) : "";
            _lblSummary.Text = string.Format("병합 대기 좌:{0} 우:{1}  (이번 {2}개 {3}){4}",
                _pendingLeft.Count, _pendingRight.Count, applied, leftToRight ? "좌→우" : "우→좌", extra);
        }

        private void SaveMerges()
        {
            if (_pendingLeft.Count == 0 && _pendingRight.Count == 0)
            {
                Info("저장할 병합 변경이 없습니다."); return;
            }
            if (IsCsvSession())
            {
                Info("CSV 테스트 세션은 저장을 지원하지 않습니다(실제 Excel 파일에서만 병합 저장)."); return;
            }
            bool backup = MessageBox.Show(this, "원본 저장 전 백업 복사본을 만들까요? (권장: 예)",
                "백업", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes;
            try
            {
                SetBusy(true);
                if (_pendingRight.Count > 0) SaveOneSide(_rightPath, _pendingRight, backup);
                if (_pendingLeft.Count > 0) SaveOneSide(_leftPath, _pendingLeft, backup);
                _pendingLeft.Clear();
                _pendingRight.Clear();
                _lblSummary.Text = "저장 완료. [비교]로 재검증을 권장합니다.";
                Info("저장이 완료되었습니다.");
            }
            catch (Exception ex)
            {
                MessageBox.Show(this,
                    "저장 실패(원본 보존):\r\n" + ex.Message +
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

        private bool IsCsvSession()
        {
            return Directory.Exists(_leftPath) || Directory.Exists(_rightPath);
        }

        private void SaveOneSide(string path, Dictionary<string, object> pending, bool backup)
        {
            List<MergeItem> items = new List<MergeItem>();
            foreach (KeyValuePair<string, object> kv in pending)
            {
                int lastColon = kv.Key.LastIndexOf(':');
                int prevColon = kv.Key.LastIndexOf(':', lastColon - 1);
                string sheet = kv.Key.Substring(0, prevColon);
                int row = int.Parse(kv.Key.Substring(prevColon + 1, lastColon - prevColon - 1));
                int col = int.Parse(kv.Key.Substring(lastColon + 1));
                items.Add(new MergeItem(sheet, row, col, kv.Value));
            }
            // COM 쓰기도 STA 필요 → 저장은 동기(짧음)로 STA 스레드에서.
            Exception error = null;
            Thread t = StaTask.Run<bool>(
                delegate
                {
                    using (MergeEngine me = new MergeEngine())
                    {
                        string backupPath;
                        me.ApplyAndSave(path, items, backup, out backupPath);
                    }
                    return true;
                },
                delegate(bool ok, Exception err) { error = err; });
            t.Join();
            if (error != null) throw error;
        }

        // ================================================================ N-way / DnD / 온보딩
        private void OpenNWayDialog()
        {
            using (NWayForm f = new NWayForm(_settings)) { f.ShowDialog(this); }
        }

        private void OnDragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy;
        }

        private void OnDragDrop(object sender, DragEventArgs e)
        {
            string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
            if (files == null || files.Length == 0) return;
            if (files.Length >= 2)
            {
                _leftPath = files[0]; _lblLeftPath.Text = files[0];
                _rightPath = files[1]; _lblRightPath.Text = files[1];
                StartCompareExcel();
            }
            else
            {
                if (string.IsNullOrEmpty(_leftPath)) { _leftPath = files[0]; _lblLeftPath.Text = files[0]; }
                else { _rightPath = files[0]; _lblRightPath.Text = files[0]; }
            }
        }

        private void ShowOnboardingIfNeeded()
        {
            if (_settings.HideOnboarding) return;
            Shown += delegate
            {
                using (Form f = new Form())
                {
                    f.Text = "환영합니다 — Excel Diff / Merge";
                    f.Width = 600; f.Height = 380;
                    f.StartPosition = FormStartPosition.CenterParent;
                    f.FormBorderStyle = FormBorderStyle.FixedDialog;
                    f.MaximizeBox = false; f.MinimizeBox = false;
                    Label lbl = new Label();
                    lbl.Dock = DockStyle.Fill; lbl.Padding = new Padding(16);
                    lbl.Text =
                        "사용법\r\n\r\n"
                      + "1) [좌측 열기]/[우측 열기] 로 파일 2개 선택 (또는 드래그앤드롭)\r\n"
                      + "2) [비교] → 좌우 병렬로 차이 표시\r\n"
                      + "     노랑=변경 초록=추가 빨강=삭제 파랑=병합됨 회색=정렬빈칸\r\n"
                      + "3) 정렬 방식 선택: 자동정렬(행 삽입/삭제 인지) / 키 컬럼 / 좌표\r\n"
                      + "4) F7/F8 로 다음/이전 차이. 우측 마커바 클릭으로 점프\r\n"
                      + "5) 셀 선택 → 하단에 좌/우 값·수식 표시\r\n"
                      + "6) [좌→우]/[우→좌] 병합 후 [저장](Ctrl+S)\r\n\r\n"
                      + "* Excel 없이 테스트: [CSV폴더비교] 로 CSV 폴더 2개 비교\r\n"
                      + "* DRM 파일은 Excel 이 복호화하여 읽습니다(우회 없음).";
                    f.Controls.Add(lbl);
                    CheckBox chk = new CheckBox();
                    chk.Text = "다시 보지 않기"; chk.Dock = DockStyle.Bottom; chk.Height = 30;
                    chk.Padding = new Padding(16, 0, 0, 0);
                    f.Controls.Add(chk);
                    Button ok = new Button();
                    ok.Text = "시작"; ok.Dock = DockStyle.Bottom; ok.Height = 34;
                    ok.Click += delegate { f.Close(); };
                    f.Controls.Add(ok);
                    f.ShowDialog(this);
                    if (chk.Checked) { _settings.HideOnboarding = true; _settings.Save(); }
                }
            };
        }

        // ================================================================ 유틸
        private void Info(string msg)
        {
            MessageBox.Show(this, msg, "안내", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private static string CellKey(string sheet, int row, int col)
        {
            return sheet + ":" + row + ":" + col;
        }

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

        /// <summary>열 문자(A, B, …, AA) → 1-based 열 번호. 실패 시 -1.</summary>
        public static int ParseColLetter(string text)
        {
            if (string.IsNullOrEmpty(text)) return -1;
            text = text.Trim().ToUpperInvariant();
            int col = 0;
            foreach (char ch in text)
            {
                if (ch < 'A' || ch > 'Z') return -1;
                col = col * 26 + (ch - 'A' + 1);
            }
            return col > 0 ? col : -1;
        }
    }
}
