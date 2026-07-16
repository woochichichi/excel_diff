using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace ExcelDiffMerge
{
    /// <summary>
    /// N-way 취합 뷰 (재설계 — 메인 화면과 같은 좌우 비교 방식).
    ///  - base 를 왼쪽에 고정, 오른쪽에는 선택한 버전 1개를 표시하는 2-way 좌우 비교 화면.
    ///  - 버전을 전환하며 하나씩 검토하고, 채택한 값은 base(왼쪽)에 누적, 마지막에 base 파일에 한 번에 저장.
    ///  - pairwise diff(base vs 각 버전)는 DiffEngine/RowAligner 로 비교 시 1회씩 계산해 캐시(버전 전환 시 재계산 없음).
    ///  - 충돌(여러 버전이 같은 셀을 다르게 변경) 좌표는 NWayDiffEngine.Compare 결과로 구해 주황 강조.
    ///  - 그리드는 가상모드(VirtualMode) + 더블버퍼링 서브클래스(MainForm 패턴 자체 재구현).
    /// </summary>
    public sealed class NWayForm : Form
    {
        // 색상 — MainForm 과 동일 팔레트.
        private static readonly Color ColChanged = Color.FromArgb(255, 243, 160); // 노랑 = 변경
        private static readonly Color ColAdded = Color.FromArgb(186, 238, 186);   // 초록 = 추가
        private static readonly Color ColDeleted = Color.FromArgb(250, 186, 186); // 빨강 = 삭제
        private static readonly Color ColAdopted = Color.FromArgb(186, 212, 255); // 파랑 = 채택됨
        private static readonly Color ColGap = Color.FromArgb(238, 238, 238);     // 반대편 없는 셀
        private static readonly Color ColConflict = Color.FromArgb(255, 200, 140);// 주황 = 충돌

        private readonly Settings _settings;

        // ===== UI =====
        private ToolStrip _tool;
        private Panel _filePanel;
        private TextBox _txtBase;
        private ListBox _lstVersions;
        private ComboBox _cboVersion;
        private TabControl _tabs;
        private DataGridView _gridLeft;
        private DataGridView _gridRight;
        private MarkerBar _marker;
        private Panel _adoptStrip;
        private Button _btnAdopt;
        private Button _btnSaveBase;
        private Label _lblLeftHdr;
        private Label _lblRightHdr;
        private Label _help;
        private ToolStripStatusLabel _lblSummary;
        private ToolStripProgressBar _progress;

        // ===== 상태 =====
        private string _basePath;
        private readonly List<string> _versionPaths = new List<string>();

        private List<DiffResult> _versionDiffs = new List<DiffResult>(); // base vs 각 버전(캐시)
        private NWayResult _nway;                                        // 충돌 좌표용
        private readonly Dictionary<string, NWayCell> _conflictCells =   // "sheet:baseRow:col" → 충돌 셀
            new Dictionary<string, NWayCell>();

        private int _curVersion = -1;
        private string _curSheetName;
        private DiffResult _curDiff;   // 선택 버전의 pairwise diff
        private SheetDiff _curSheet;   // 선택 버전 + 현재 시트

        // 표시행 매핑(가상 그리드).
        private int[] _rowMap = new int[0];
        private Dictionary<int, int> _rowsIdxToDisplay = new Dictionary<int, int>();
        private int _colStart = 1;
        private int _colCount;

        // base 채택 대기(누적). key "sheet:baseRow:col" → 채택값 + 출처 버전.
        private readonly Dictionary<string, AdoptEntry> _adopt = new Dictionary<string, AdoptEntry>();
        // Ctrl+Z 되돌리기(동작 단위).
        private readonly Stack<AdoptUndo> _undo = new Stack<AdoptUndo>();

        private bool _syncing;
        private bool _busy;
        private int _navIndex = -1;

        /// <summary>채택된 값 한 개 + 어느 버전에서 채택했는지(교체 확인용).</summary>
        private sealed class AdoptEntry
        {
            public object Value;
            public int Version;
            public AdoptEntry(object value, int version) { Value = value; Version = version; }
        }

        /// <summary>채택/교체/취소 한 번을 되돌릴 언두 단위.</summary>
        private sealed class AdoptUndo
        {
            public string Key;
            public bool Existed;
            public AdoptEntry Prev;
        }

        /// <summary>로드/비교 결과 묶음(STA 워커 → UI 마샬링).</summary>
        private sealed class CompareData
        {
            public WorkbookData Base;
            public List<WorkbookData> Versions;
            public List<DiffResult> Diffs;
            public NWayResult Nway;
        }

        /// <summary>더블버퍼링 DataGridView(MainForm.BufferedGrid 와 동일 패턴을 여기 자체 정의).</summary>
        private sealed class BufferedGrid : DataGridView
        {
            public BufferedGrid()
            {
                DoubleBuffered = true;
                SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            }
        }

        public NWayForm(Settings settings)
        {
            _settings = settings;
            BuildUi();
        }

        // ================================================================ UI
        private void BuildUi()
        {
            Text = "N-way 취합 (base 좌 · 버전 우 — 좌우 비교)";
            Width = 1320; Height = 820;
            StartPosition = FormStartPosition.CenterParent;
            KeyPreview = true;
            KeyDown += OnKeyDown;

            // ----- 툴바(채택/저장 버튼은 툴바에서 제거) -----
            _tool = new ToolStrip();
            _tool.GripStyle = ToolStripGripStyle.Hidden;
            _tool.Dock = DockStyle.Top;
            _tool.Font = new Font("Segoe UI", 11F, FontStyle.Regular);
            _tool.ImageScalingSize = new Size(1, 1);
            _tool.Renderer = new ToolStripProfessionalRenderer();
            _tool.Padding = new Padding(3, 3, 3, 3);
            AddBtn(_tool, "① base 선택", delegate { PickBase(); });
            AddBtn(_tool, "② 버전 추가", delegate { AddVersion(); });
            AddBtn(_tool, "버전 제거", delegate { RemoveVersion(); });
            _tool.Items.Add(new ToolStripSeparator());
            AddBtn(_tool, "③ 비교", delegate { StartCompare(); });
            _tool.Items.Add(new ToolStripSeparator());
            AddBtn(_tool, "파일 목록 접기/펼치기", delegate { ToggleFilePanel(); });

            // ----- 파일 선택 패널(비교 후 자동 접힘) -----
            _filePanel = new Panel();
            _filePanel.Dock = DockStyle.Top;
            _filePanel.Height = 120;

            Label lb = new Label();
            lb.Text = "base:";
            lb.Location = new Point(8, 8);
            lb.AutoSize = true;
            _filePanel.Controls.Add(lb);

            _txtBase = new TextBox();
            _txtBase.Location = new Point(60, 4);
            _txtBase.Width = 1200;
            _txtBase.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            _txtBase.ReadOnly = true;
            _filePanel.Controls.Add(_txtBase);

            Label lv = new Label();
            lv.Text = "버전들:";
            lv.Location = new Point(8, 34);
            lv.AutoSize = true;
            _filePanel.Controls.Add(lv);

            _lstVersions = new ListBox();
            _lstVersions.Location = new Point(60, 32);
            _lstVersions.Width = 1200;
            _lstVersions.Height = 80;
            _lstVersions.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right;
            _filePanel.Controls.Add(_lstVersions);

            // ----- 버전 전환 바 -----
            Panel verBar = new Panel();
            verBar.Dock = DockStyle.Top;
            verBar.Height = 32;
            Label lvcap = new Label();
            lvcap.Text = "검토 버전:";
            lvcap.AutoSize = true;
            lvcap.Location = new Point(8, 8);
            verBar.Controls.Add(lvcap);
            _cboVersion = new ComboBox();
            _cboVersion.DropDownStyle = ComboBoxStyle.DropDownList;
            _cboVersion.Location = new Point(80, 4);
            _cboVersion.Width = 520;
            _cboVersion.Font = new Font("Segoe UI", 10F);
            _cboVersion.SelectedIndexChanged += delegate { OnVersionChanged(); };
            verBar.Controls.Add(_cboVersion);

            // ----- 시트 탭 -----
            _tabs = new TabControl();
            _tabs.Dock = DockStyle.Top;
            _tabs.Height = 28;
            _tabs.SelectedIndexChanged += delegate { OnSheetChanged(); };

            // ----- 좌/우 헤더(파일명) -----
            Panel hdr = new Panel();
            hdr.Dock = DockStyle.Top;
            hdr.Height = 24;
            Panel hL = new Panel(); hL.Dock = DockStyle.Left; hL.Width = 640;
            _lblLeftHdr = new Label();
            _lblLeftHdr.Text = "base (왼쪽)";
            _lblLeftHdr.Dock = DockStyle.Fill;
            _lblLeftHdr.TextAlign = ContentAlignment.MiddleLeft;
            _lblLeftHdr.AutoEllipsis = true;
            _lblLeftHdr.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            hL.Controls.Add(_lblLeftHdr);
            Panel hR = new Panel(); hR.Dock = DockStyle.Fill;
            _lblRightHdr = new Label();
            _lblRightHdr.Text = "선택한 버전 (오른쪽)";
            _lblRightHdr.Dock = DockStyle.Fill;
            _lblRightHdr.TextAlign = ContentAlignment.MiddleLeft;
            _lblRightHdr.AutoEllipsis = true;
            _lblRightHdr.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            hR.Controls.Add(_lblRightHdr);
            hdr.Controls.Add(hR);
            hdr.Controls.Add(hL);

            // ----- 좌우 그리드 + 중앙 채택 스트립 + 마커 + 좌 하단 저장바 -----
            SplitContainer split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.SplitterWidth = 4;
            _gridLeft = MakeGrid();
            _gridRight = MakeGrid();

            // 좌 그리드 하단: [base 에 저장 (N)] (우측 정렬)
            _btnSaveBase = new Button();
            _btnSaveBase.Text = "base 에 저장";
            _btnSaveBase.Dock = DockStyle.Right;
            _btnSaveBase.Width = 200;
            _btnSaveBase.Enabled = false;
            _btnSaveBase.Click += delegate { SaveAdopted(); };
            Panel bottomLeft = new Panel();
            bottomLeft.Dock = DockStyle.Bottom; bottomLeft.Height = 32;
            bottomLeft.Controls.Add(_btnSaveBase);

            split.Panel1.Controls.Add(_gridLeft);
            split.Panel1.Controls.Add(bottomLeft);

            _marker = new MarkerBar();
            _marker.Dock = DockStyle.Right;
            _marker.Width = 20;
            _marker.OnSeek = delegate(float pos) { SeekTo(pos); };

            // 중앙 세로 스트립: [←] 채택 버튼(버전 값을 base 에 채택). → 방향 없음.
            _adoptStrip = new Panel();
            _adoptStrip.Dock = DockStyle.Left;
            _adoptStrip.Width = 46;
            _adoptStrip.BackColor = SystemColors.Control;
            _btnAdopt = new Button();
            _btnAdopt.Text = "←";
            _btnAdopt.Width = 38; _btnAdopt.Height = 44;
            _btnAdopt.Font = new Font("Segoe UI", 14F, FontStyle.Bold);
            _btnAdopt.FlatStyle = FlatStyle.Standard;
            _btnAdopt.TabStop = false;
            ToolTip adoptTip = new ToolTip();
            adoptTip.SetToolTip(_btnAdopt, "선택한 오른쪽(버전) 셀 값을 base 에 채택 — 오른쪽 차이 셀 더블클릭도 동일");
            _btnAdopt.Click += delegate { AdoptCurrent(); };
            _adoptStrip.Controls.Add(_btnAdopt);
            _adoptStrip.Resize += delegate { LayoutAdoptStrip(); };

            split.Panel2.Controls.Add(_gridRight);
            split.Panel2.Controls.Add(_adoptStrip);
            split.Panel2.Controls.Add(_marker);

            // 그리드 영역을 덮는 사용법 안내(비교 성공 시 숨김).
            Panel gridHost = new Panel();
            gridHost.Dock = DockStyle.Fill;
            gridHost.Controls.Add(split);
            _help = new Label();
            _help.Dock = DockStyle.Fill;
            _help.TextAlign = ContentAlignment.MiddleLeft;
            _help.Padding = new Padding(24);
            _help.Font = new Font("Segoe UI", 10F);
            _help.BackColor = SystemColors.Window;
            _help.Text =
                "■ N-way 취합 (좌우 비교 방식)\r\n"
              + "   왼쪽에 기준 파일(base)을 고정하고, 오른쪽에 버전 1개를 표시해 메인 화면처럼 좌우로 비교합니다.\r\n"
              + "   버전을 전환하며 하나씩 검토하고, 채택한 값은 base(왼쪽)에 파랑으로 누적됩니다.\r\n\r\n"
              + "■ 사용 순서\r\n"
              + "   ① [base 선택]  — 기준이 되는 원본 파일 1개\r\n"
              + "   ② [버전 추가]  — 비교할 수정본들을 하나씩 (2개 이상 권장)\r\n"
              + "   ③ [비교]        — base 대비 각 버전을 좌우로 표시(변경 건수는 버전 콤보에 표시)\r\n"
              + "   → 오른쪽에서 원하는 값을 고르고 가운데 [←] 로 채택(또는 오른쪽 차이 셀 더블클릭)\r\n"
              + "   → 다 고른 뒤 [base 에 저장] 으로 base 파일에 한 번에 저장\r\n\r\n"
              + "   색상: 노랑=변경  초록=추가  빨강=삭제  파랑=채택됨  주황=충돌(다른 버전도 다르게 변경)\r\n"
              + "   F7/F8=다음/이전 차이,  Ctrl+Z=마지막 채택 취소";
            gridHost.Controls.Add(_help);
            _help.BringToFront();

            Load += delegate { try { split.SplitterDistance = split.Width / 2; } catch { } LayoutAdoptStrip(); };

            // ----- 상태바 -----
            StatusStrip status = new StatusStrip();
            _lblSummary = new ToolStripStatusLabel("① base 선택 → ② 버전 추가 → ③ 비교 하세요.");
            _lblSummary.Spring = true;
            _lblSummary.TextAlign = ContentAlignment.MiddleLeft;
            _progress = new ToolStripProgressBar();
            _progress.Style = ProgressBarStyle.Marquee;
            _progress.Visible = false;
            status.Items.Add(_lblSummary);
            status.Items.Add(_progress);
            AddStatusLegend(status, ColChanged, "변경");
            AddStatusLegend(status, ColAdded, "추가");
            AddStatusLegend(status, ColDeleted, "삭제");
            AddStatusLegend(status, ColAdopted, "채택됨");
            AddStatusLegend(status, ColConflict, "충돌");

            // 도킹 z-order: Fill 먼저, Top 은 안쪽(탭)→바깥쪽(툴바) 순, 끝에 Bottom.
            Controls.Add(gridHost);   // Fill
            Controls.Add(hdr);        // Top (그리드 바로 위)
            Controls.Add(_tabs);      // Top
            Controls.Add(verBar);     // Top
            Controls.Add(_filePanel); // Top
            Controls.Add(_tool);      // Top (최상단)
            Controls.Add(status);     // Bottom
        }

        private static ToolStripButton AddBtn(ToolStrip t, string text, EventHandler h)
        {
            ToolStripButton b = new ToolStripButton(text);
            b.DisplayStyle = ToolStripItemDisplayStyle.Text;
            b.Padding = new Padding(10, 7, 10, 7);
            b.Margin = new Padding(3, 2, 3, 2);
            b.BackColor = Color.FromArgb(235, 240, 250);
            b.Click += h;
            t.Items.Add(b);
            return b;
        }

        private static void AddStatusLegend(StatusStrip status, Color color, string text)
        {
            ToolStripStatusLabel sw = new ToolStripStatusLabel("  ");
            sw.BackColor = color;
            sw.BorderSides = ToolStripStatusLabelBorderSides.All;
            sw.BorderStyle = Border3DStyle.Flat;
            sw.Margin = new Padding(4, 3, 0, 3);
            status.Items.Add(sw);
            ToolStripStatusLabel lb = new ToolStripStatusLabel(text);
            lb.Margin = new Padding(2, 0, 6, 0);
            status.Items.Add(lb);
        }

        private DataGridView MakeGrid()
        {
            DataGridView g = new BufferedGrid();
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
            g.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText;
            g.ShowCellToolTips = true;
            g.CellValueNeeded += Grid_CellValueNeeded;
            g.CellFormatting += Grid_CellFormatting;
            g.CellToolTipTextNeeded += Grid_CellToolTipTextNeeded;
            g.Scroll += Grid_Scroll;
            g.SelectionChanged += Grid_SelectionChanged;
            g.RowPostPaint += Grid_RowPostPaint;
            g.CellDoubleClick += Grid_CellDoubleClick;
            return g;
        }

        private void LayoutAdoptStrip()
        {
            if (_adoptStrip == null || _btnAdopt == null) return;
            int w = _adoptStrip.ClientSize.Width;
            int x = Math.Max(2, (w - _btnAdopt.Width) / 2);
            int y = Math.Max(4, (_adoptStrip.ClientSize.Height - _btnAdopt.Height) / 2);
            _btnAdopt.Left = x; _btnAdopt.Top = y;
        }

        // ================================================================ 파일 선택
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

        private void ToggleFilePanel()
        {
            if (_filePanel == null) return;
            _filePanel.Visible = !_filePanel.Visible;
        }

        // ================================================================ 비교(로드 + pairwise diff + N-way 충돌)
        private void StartCompare()
        {
            if (string.IsNullOrEmpty(_basePath) || _versionPaths.Count == 0)
            {
                MessageBox.Show(this, "base 파일과 버전 1개 이상을 지정하세요.", "안내",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (_busy) return;
            _busy = true;
            _progress.Visible = true;
            _lblSummary.Text = "로드/비교 중…";

            string basePath = _basePath;
            List<string> versionPaths = new List<string>(_versionPaths);

            StaTask.Run<CompareData>(
                delegate
                {
                    // 파일마다 '자기 리더'로 열고 닫는다(B3, MainForm.LoadWorkbookSafe 패턴).
                    WorkbookData baseWb = LoadWorkbookSafe(basePath);
                    List<WorkbookData> versions = new List<WorkbookData>();
                    foreach (string p in versionPaths) versions.Add(LoadWorkbookSafe(p));

                    // pairwise diff(base vs 각 버전) — 메인과 동일 자동정렬(LCS) 모드.
                    List<DiffResult> diffs = new List<DiffResult>();
                    foreach (WorkbookData v in versions)
                        diffs.Add(DiffEngine.Compare(baseWb, v, AlignMode.Auto, -1));

                    // 충돌 좌표는 NWayDiffEngine 로(수정 금지, 그대로 호출).
                    NWayResult nw = NWayDiffEngine.Compare(baseWb, versions);

                    CompareData cd = new CompareData();
                    cd.Base = baseWb; cd.Versions = versions; cd.Diffs = diffs; cd.Nway = nw;
                    return cd;
                },
                delegate(CompareData cd, Exception err)
                {
                    BeginInvoke((MethodInvoker)delegate { OnLoaded(cd, err); });
                });
        }

        private static WorkbookData LoadWorkbookSafe(string path)
        {
            using (IWorkbookReader reader = MakeReader(path))
            {
                return reader.LoadWorkbook(path);
            }
        }

        private static IWorkbookReader MakeReader(string path)
        {
            if (Directory.Exists(path)) return new CsvWorkbookReader();
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".doc" || ext == ".docx" || ext == ".docm") return new WordComReader();
            return new ExcelComReader();
        }

        private void OnLoaded(CompareData cd, Exception err)
        {
            _busy = false;
            _progress.Visible = false;
            if (err != null)
            {
                _lblSummary.Text = "오류: " + err.Message;
                MessageBox.Show(this, err.Message, "비교 실패",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            _versionDiffs = cd.Diffs;
            _nway = cd.Nway;

            // 새 비교 → 채택/언두 초기화.
            _adopt.Clear();
            _undo.Clear();

            BuildConflictIndex();

            if (_help != null) _help.Visible = false;
            _lblLeftHdr.Text = "base (왼쪽): " + Path.GetFileName(_basePath);

            PopulateVersionCombo();
            BuildTabs();

            // 비교 후 파일 선택 패널 자동 접기 → 그리드 공간 확보.
            if (_filePanel != null) _filePanel.Visible = false;

            _curVersion = -1;
            if (_cboVersion.Items.Count > 0) _cboVersion.SelectedIndex = 0; // → OnVersionChanged
        }

        private void BuildConflictIndex()
        {
            _conflictCells.Clear();
            if (_nway == null) return;
            foreach (NWaySheetDiff s in _nway.Sheets)
                foreach (NWayCell c in s.Cells)
                    if (c.Conflict)
                        _conflictCells[CellKey(s.Name, c.Row, c.Col)] = c;
        }

        private void PopulateVersionCombo()
        {
            _cboVersion.Items.Clear();
            for (int i = 0; i < _versionDiffs.Count; i++)
            {
                DiffResult d = _versionDiffs[i];
                int changes = d.TotalChanged + d.TotalAdded + d.TotalDeleted;
                _cboVersion.Items.Add("v" + (i + 1) + ": " + Path.GetFileName(_versionPaths[i])
                    + "  (변경 " + changes + ")");
            }
        }

        // ================================================================ 버전/시트 전환
        private void OnVersionChanged()
        {
            int idx = _cboVersion.SelectedIndex;
            if (idx < 0 || idx >= _versionDiffs.Count) return;
            _curVersion = idx;
            _curDiff = _versionDiffs[idx];
            _lblRightHdr.Text = "v" + (idx + 1) + " (오른쪽): " + Path.GetFileName(_versionPaths[idx]);
            UpdateTabMarkers();
            // 현재 시트를 새 버전 기준으로 다시 그림(시트 유지).
            OnSheetChanged();
            UpdateSummary();
        }

        private void BuildTabs()
        {
            _tabs.TabPages.Clear();
            // 시트 union(base + 모든 버전). N-way 결과의 시트 순서를 사용.
            List<string> names = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (_nway != null)
                foreach (NWaySheetDiff s in _nway.Sheets)
                    if (seen.Add(s.Name)) names.Add(s.Name);
            // 안전망: 혹시 빠진 시트가 있으면 diff 에서 보충.
            foreach (DiffResult d in _versionDiffs)
                foreach (SheetDiff sd in d.Sheets)
                    if (seen.Add(sd.Name)) names.Add(sd.Name);

            foreach (string name in names)
            {
                TabPage tp = new TabPage(name);
                tp.Tag = name;
                _tabs.TabPages.Add(tp);
            }
        }

        /// <summary>선택 버전 기준으로 변경 있는 시트에 ● 표시.</summary>
        private void UpdateTabMarkers()
        {
            if (_curDiff == null) return;
            foreach (TabPage tp in _tabs.TabPages)
            {
                string name = (string)tp.Tag;
                SheetDiff sd = FindSheetDiff(_curDiff, name);
                int n = sd != null ? (sd.ChangedCount + sd.AddedCount + sd.DeletedCount) : 0;
                tp.Text = name + (n > 0 ? " ●(" + n + ")" : "");
            }
        }

        private static SheetDiff FindSheetDiff(DiffResult d, string name)
        {
            if (d == null) return null;
            foreach (SheetDiff sd in d.Sheets)
                if (string.Equals(sd.Name, name, StringComparison.OrdinalIgnoreCase)) return sd;
            return null;
        }

        private void OnSheetChanged()
        {
            if (_tabs.SelectedTab != null) _curSheetName = (string)_tabs.SelectedTab.Tag;
            if (_curSheetName == null && _tabs.TabPages.Count > 0)
            {
                _tabs.SelectedIndex = 0;
                _curSheetName = (string)_tabs.SelectedTab.Tag;
            }
            _curSheet = FindSheetDiff(_curDiff, _curSheetName);
            _navIndex = -1;

            if (_curSheet != null && _curSheet.ColCount > 0)
            {
                _colStart = _curSheet.MinCol;
                _colCount = _curSheet.ColCount;
            }
            else { _colStart = 1; _colCount = 0; }

            BuildColumns(_gridLeft);
            BuildColumns(_gridRight);
            RefreshGridRows();
            UpdateSummary();
        }

        private void BuildColumns(DataGridView g)
        {
            g.RowCount = 0;
            g.Columns.Clear();
            if (_colCount <= 0) return;
            DataGridViewColumn[] cols = new DataGridViewColumn[_colCount];
            int width = _colCount == 1 ? 640 : 110;
            for (int i = 0; i < _colCount; i++)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.HeaderText = _colCount == 1 ? "본문" : MainForm.ColLetter(_colStart + i);
                col.Width = width;
                col.SortMode = DataGridViewColumnSortMode.NotSortable;
                cols[i] = col;
            }
            g.Columns.AddRange(cols);
        }

        private void RefreshGridRows()
        {
            if (_curSheet == null)
            {
                _rowMap = new int[0];
                _rowsIdxToDisplay = new Dictionary<int, int>();
                SetRowCount(_gridLeft, 0);
                SetRowCount(_gridRight, 0);
                if (_marker != null) _marker.SetMarks(null);
                return;
            }
            List<int> rows = new List<int>();
            for (int i = 0; i < _curSheet.Rows.Count; i++) rows.Add(i);
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

            if (isLeft)
            {
                if (dr.LeftRow < 0) { e.Value = ""; return; }
                // base 유효값: 채택 대기가 있으면 그 값 우선(파랑 표시와 일치).
                if (_adopt.Count > 0)
                {
                    AdoptEntry ae;
                    if (_adopt.TryGetValue(CellKey(_curSheetName, dr.LeftRow, absCol), out ae))
                    { e.Value = DiffEngine.ToText(ae.Value); return; }
                }
                object bv = _curSheet.Left != null ? _curSheet.Left.GetValueAbs(dr.LeftRow, absCol) : null;
                e.Value = DiffEngine.ToText(bv);
            }
            else
            {
                if (dr.RightRow < 0) { e.Value = ""; return; }
                object vv = _curSheet.Right != null ? _curSheet.Right.GetValueAbs(dr.RightRow, absCol) : null;
                e.Value = DiffEngine.ToText(vv);
            }
        }

        private void Grid_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (_curSheet == null || e.RowIndex < 0 || e.RowIndex >= _rowMap.Length) return;
            bool isLeft = (sender == _gridLeft);
            DiffRow dr = _curSheet.Rows[_rowMap[e.RowIndex]];
            int absCol = _colStart + e.ColumnIndex;
            CellStatus st = dr.StatusAt(absCol);

            if (isLeft)
            {
                if (dr.LeftRow < 0) { e.CellStyle.BackColor = ColGap; return; }
                string key = CellKey(_curSheetName, dr.LeftRow, absCol);
                if (_adopt.Count > 0 && _adopt.ContainsKey(key)) { e.CellStyle.BackColor = ColAdopted; return; }
                if (_conflictCells.ContainsKey(key)) { e.CellStyle.BackColor = ColConflict; return; }
                switch (st)
                {
                    case CellStatus.Changed: e.CellStyle.BackColor = ColChanged; break;
                    case CellStatus.Deleted: e.CellStyle.BackColor = ColDeleted; break;
                    case CellStatus.Added: e.CellStyle.BackColor = ColGap; break;
                }
            }
            else
            {
                if (dr.RightRow < 0) { e.CellStyle.BackColor = ColGap; return; }
                // 충돌 오버레이(변경 셀에 한해 주황): 다른 버전도 이 base 셀을 다르게 변경.
                if (st != CellStatus.Same && dr.LeftRow >= 0
                    && _conflictCells.ContainsKey(CellKey(_curSheetName, dr.LeftRow, absCol)))
                { e.CellStyle.BackColor = ColConflict; return; }
                switch (st)
                {
                    case CellStatus.Changed: e.CellStyle.BackColor = ColChanged; break;
                    case CellStatus.Added: e.CellStyle.BackColor = ColAdded; break;
                    case CellStatus.Deleted: e.CellStyle.BackColor = ColGap; break;
                }
            }
        }

        private void Grid_CellToolTipTextNeeded(object sender, DataGridViewCellToolTipTextNeededEventArgs e)
        {
            if (_curSheet == null || e.RowIndex < 0 || e.RowIndex >= _rowMap.Length || e.ColumnIndex < 0) return;
            DiffRow dr = _curSheet.Rows[_rowMap[e.RowIndex]];
            if (dr.LeftRow < 0) return;
            int absCol = _colStart + e.ColumnIndex;
            NWayCell nc;
            if (_conflictCells.TryGetValue(CellKey(_curSheetName, dr.LeftRow, absCol), out nc))
                e.ToolTipText = BuildConflictTip(nc);
        }

        private string BuildConflictTip(NWayCell nc)
        {
            string s = "여러 버전이 이 셀을 다르게 변경했습니다.\r\nbase: " + DiffEngine.ToText(nc.BaseValue);
            foreach (KeyValuePair<int, object> kv in nc.Changes)
                s += "\r\nv" + (kv.Key + 1) + ": " + DiffEngine.ToText(kv.Value);
            s += "\r\n→ 다른 버전도 확인 후 채택하세요.";
            return s;
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
                if (src.CurrentCell != null && dst.RowCount > 0 && dst.ColumnCount > 0)
                {
                    int r = src.CurrentCell.RowIndex, c = src.CurrentCell.ColumnIndex;
                    if (r >= 0 && r < dst.RowCount && c >= 0 && c < dst.ColumnCount)
                    {
                        try
                        {
                            dst.CurrentCell = dst.Rows[r].Cells[c];
                            if (!dst.Rows[r].Cells[c].Selected)
                            {
                                dst.ClearSelection();
                                dst.Rows[r].Cells[c].Selected = true;
                            }
                        }
                        catch { }
                    }
                }
            }
            catch { }
            finally { _syncing = false; }
        }

        // ================================================================ 채택
        private void Grid_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            AdoptAt(e.RowIndex, e.ColumnIndex);
        }

        /// <summary>현재 선택 셀(오른쪽 우선)을 채택.</summary>
        private void AdoptCurrent()
        {
            DataGridView g = _gridRight.CurrentCell != null ? _gridRight : _gridLeft;
            if (g.CurrentCell == null) return;
            AdoptAt(g.CurrentCell.RowIndex, g.CurrentCell.ColumnIndex);
        }

        /// <summary>표시행/열 좌표의 '오른쪽(버전) 값'을 base 채택 대기로 등록.</summary>
        private void AdoptAt(int displayRow, int displayCol)
        {
            if (_curSheet == null || displayRow < 0 || displayRow >= _rowMap.Length || displayCol < 0) return;

            // base 에 없는 시트는 저장 불가 → 채택 차단(B2, InBase 로직 유지).
            if (_curSheet.Left == null)
            {
                MessageBox.Show(this,
                    "이 시트 '" + _curSheetName + "' 은(는) base 파일에 없는 시트라 base 에 채택/저장할 수 없습니다.\r\n"
                    + "(버전 파일에만 존재하는 시트입니다.)",
                    "채택 불가", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DiffRow dr = _curSheet.Rows[_rowMap[displayRow]];
            int absCol = _colStart + displayCol;

            if (dr.LeftRow < 0 || dr.RightRow < 0)
            {
                _lblSummary.Text = "이 셀은 행 추가/삭제라 base 채택을 지원하지 않습니다.";
                return;
            }
            if (dr.StatusAt(absCol) == CellStatus.Same)
            {
                _lblSummary.Text = "동일한 셀이라 채택할 것이 없습니다.";
                return;
            }

            object versionVal = _curSheet.Right != null ? _curSheet.Right.GetValueAbs(dr.RightRow, absCol) : null;
            string key = CellKey(_curSheetName, dr.LeftRow, absCol);

            // 이미 다른 버전에서 채택한 셀을 다른 값으로 다시 채택 → 교체 확인.
            AdoptEntry existing;
            if (_adopt.TryGetValue(key, out existing))
            {
                if (existing.Version != _curVersion && !DiffEngine.ValueEquals(existing.Value, versionVal))
                {
                    DialogResult a = MessageBox.Show(this,
                        "v" + (existing.Version + 1) + " 값으로 채택한 셀입니다.\r\n"
                        + "v" + (_curVersion + 1) + " 값('" + DiffEngine.ToText(versionVal) + "')으로 교체할까요?",
                        "채택 교체", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (a != DialogResult.Yes) return;
                }
            }

            ApplyAdopt(key, new AdoptEntry(versionVal, _curVersion));
            _gridLeft.Invalidate();
            _gridRight.Invalidate();
            _lblSummary.Text = "채택: " + MainForm.ColLetter(absCol) + dr.LeftRow
                + " ← v" + (_curVersion + 1) + " '" + DiffEngine.ToText(versionVal) + "'  (채택 대기 " + _adopt.Count + ")";
            UpdateSaveButton();
        }

        /// <summary>채택 적용 + 언두 1단위 기록.</summary>
        private void ApplyAdopt(string key, AdoptEntry entry)
        {
            AdoptUndo u = new AdoptUndo();
            u.Key = key;
            AdoptEntry prev;
            u.Existed = _adopt.TryGetValue(key, out prev);
            u.Prev = u.Existed ? prev : null;
            _undo.Push(u);
            _adopt[key] = entry;
        }

        private void UndoLastAdopt()
        {
            if (_undo.Count == 0) { _lblSummary.Text = "되돌릴 채택이 없습니다."; return; }
            AdoptUndo u = _undo.Pop();
            if (u.Existed) _adopt[u.Key] = u.Prev;
            else _adopt.Remove(u.Key);
            _gridLeft.Invalidate();
            _gridRight.Invalidate();
            _lblSummary.Text = "마지막 채택 취소  (채택 대기 " + _adopt.Count + ")";
            UpdateSaveButton();
        }

        // ================================================================ 네비게이션(현재 버전 기준)
        private void OnKeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F7) { NavigateDiff(1); e.Handled = true; }
            else if (e.KeyCode == Keys.F8) { NavigateDiff(-1); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.Z) { UndoLastAdopt(); e.Handled = true; e.SuppressKeyPress = true; }
        }

        private void NavigateDiff(int dir)
        {
            if (_curSheet == null || _curSheet.Nav.Count == 0) return;
            int ni = _navIndex + dir;
            if (ni < 0) ni = _curSheet.Nav.Count - 1;
            if (ni >= _curSheet.Nav.Count) ni = 0;
            _navIndex = ni;
            SelectCell(_curSheet.Nav[ni]);
        }

        private void SelectCell(DiffNav nv)
        {
            if (_curSheet == null) return;
            int displayRow;
            if (!_rowsIdxToDisplay.TryGetValue(nv.RowIndex, out displayRow)) return;
            int displayCol = nv.Col - _colStart;
            if (displayCol < 0 || displayCol >= _colCount) return;
            try
            {
                _gridRight.CurrentCell = _gridRight.Rows[displayRow].Cells[displayCol];
                _gridRight.FirstDisplayedScrollingRowIndex = Math.Max(0, displayRow - 3);
                _gridLeft.FirstDisplayedScrollingRowIndex = Math.Max(0, displayRow - 3);
            }
            catch { }
            _lblSummary.Text = string.Format("차이 {0}/{1} — {2}{3}  (v{4})",
                _navIndex + 1, _curSheet.Nav.Count, MainForm.ColLetter(nv.Col),
                RowLabel(nv.RowIndex), _curVersion + 1);
        }

        private string RowLabel(int rowsIdx)
        {
            DiffRow dr = _curSheet.Rows[rowsIdx];
            int r = dr.LeftRow >= 0 ? dr.LeftRow : dr.RightRow;
            return r >= 0 ? r.ToString() : "-";
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

        // ================================================================ 요약/저장
        private void UpdateSummary()
        {
            if (_curDiff == null) { UpdateSaveButton(); return; }
            int curChanges = _curDiff.TotalChanged + _curDiff.TotalAdded + _curDiff.TotalDeleted;
            int conflicts = _nway != null ? _nway.TotalConflict : 0;
            _lblSummary.Text = string.Format("v{0} 검토 중 — 변경 {1}건, 채택 대기 {2}건, 충돌 {3}건",
                _curVersion + 1, curChanges, _adopt.Count, conflicts);
            UpdateSaveButton();
        }

        private void UpdateSaveButton()
        {
            if (_btnSaveBase == null) return;
            int n = _adopt.Count;
            _btnSaveBase.Enabled = n > 0;
            _btnSaveBase.Text = n > 0 ? ("base 에 저장 (" + n + ")") : "base 에 저장";
        }

        private void SaveAdopted()
        {
            if (_adopt.Count == 0)
            {
                MessageBox.Show(this, "채택한 값이 없습니다.", "안내",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (Directory.Exists(_basePath))
            {
                MessageBox.Show(this, "CSV 폴더(테스트) base 는 저장을 지원하지 않습니다(실제 Excel 파일에서만 저장).",
                    "저장 불가", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            DialogResult dr = MessageBox.Show(this,
                "채택한 " + _adopt.Count + "개 셀을 base 파일에 저장합니다:\r\n" + _basePath
                + "\r\n\r\n[예] 백업 복사본을 만든 뒤 저장\r\n[아니오] 백업 없이 저장\r\n[취소] 저장 안 함"
                + "\r\n\r\n※ base 파일이 Excel 에서 열려 있으면 저장이 실패합니다. 먼저 닫아 주세요.",
                "base 에 저장", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.Cancel) return;
            bool backup = (dr == DialogResult.Yes);

            List<MergeItem> items = new List<MergeItem>();
            foreach (KeyValuePair<string, AdoptEntry> kv in _adopt)
            {
                int lastColon = kv.Key.LastIndexOf(':');
                int prevColon = kv.Key.LastIndexOf(':', lastColon - 1);
                string sheet = kv.Key.Substring(0, prevColon);
                int row = int.Parse(kv.Key.Substring(prevColon + 1, lastColon - prevColon - 1));
                int col = int.Parse(kv.Key.Substring(lastColon + 1));
                items.Add(new MergeItem(sheet, row, col, kv.Value.Value));
            }

            try
            {
                List<string> skippedSheets = null;
                Exception error = null;
                Thread t = StaTask.Run<bool>(
                    delegate
                    {
                        using (MergeEngine me = new MergeEngine())
                        {
                            string backupPath;
                            List<string> skipped;
                            me.ApplyAndSave(_basePath, items, backup, out backupPath, out skipped);
                            skippedSheets = skipped;
                        }
                        return true;
                    },
                    delegate(bool ok, Exception err) { error = err; });
                t.Join();
                if (error != null) throw error;

                _adopt.Clear();
                _undo.Clear();
                _gridLeft.Invalidate();
                _gridRight.Invalidate();
                UpdateSaveButton();

                string msg = "base 파일에 저장 완료.";
                if (skippedSheets != null && skippedSheets.Count > 0)
                    msg += "\r\n\r\n※ 저장 불가 시트 " + skippedSheets.Count
                         + "개(base에 없음): " + string.Join(", ", skippedSheets.ToArray())
                         + "\r\n  → 해당 시트의 셀은 저장되지 않았습니다. 나머지는 정상 저장되었습니다.";
                MessageBox.Show(this, msg, "완료",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "저장 실패(원본 보존):\r\n" + ex.Message, "저장 실패",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // ================================================================ 유틸
        private static string CellKey(string sheet, int row, int col)
        {
            return sheet + ":" + row + ":" + col;
        }
    }
}
