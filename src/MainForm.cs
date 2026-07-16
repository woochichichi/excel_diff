using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
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
        private ToolStrip _tool2;
        private ToolStripButton _btnPrev;
        private ToolStripButton _btnNext;
        private ToolStripButton _btnMergeLR;
        private ToolStripButton _btnMergeRL;
        private ToolStripButton _btnSave;
        private ToolStripButton _btnMergeAllLR;
        private ToolStripButton _btnMergeAllRL;
        private ToolStripButton _btnChangeList;
        private ToolStripComboBox _cboAlign;
        private ToolStripTextBox _txtKeyCol;
        private Panel _changePanel;
        private ListView _changeList;
        private Label _changeHeader;
        private DataGridView _gridLeft;
        private DataGridView _gridRight;
        private MarkerBar _marker;
        private TabControl _tabs;
        private Label _lblLeftPath;
        private Label _lblRightPath;
        private ToolStripStatusLabel _lblSummary;
        private ToolStripProgressBar _progress;
        private ToolStripButton _btnChangesOnly;
        private RichTextBox _detail;

        // 최근 비교 드롭다운 / 폰트 / 병합취소 등 추가 UI (개선)
        private ToolStripDropDownButton _btnRecent;
        private ToolStripButton _btnFontUp;
        private ToolStripButton _btnFontDown;
        private ToolStripButton _btnCancelMerge;
        private ToolStripButton _btnAutoFit;
        private TextBox _txtSearch;      // 변경목록 검색(U4)
        private ComboBox _cboKind;       // 변경목록 종류 필터(U4)
        private ContextMenuStrip _cellMenu;   // 그리드 셀 우클릭(U8)
        private ToolStripMenuItem _cellMenuCancel;
        private DataGridView _ctxGrid;   // 마지막 우클릭 그리드/셀
        private int _ctxRow = -1;
        private int _ctxCol = -1;
        private float _gridFontSize = 9f;
        private Font _gridFont;                       // 현재 그리드 폰트(교체 시 이전 것 Dispose)

        // 파일별 저장 버튼(U2) / 경로 툴팁
        private Button _btnSaveLeftFile;
        private Button _btnSaveRightFile;
        private ToolTip _pathTip;

        // 상세 패널 중복 갱신 방지(같은 셀 재선택 시 재계산 생략 — 속도)
        private int _lastDetailRow = -1;
        private int _lastDetailCol = -1;

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
        private bool _isPreview;   // 파일 열기 직후 미리보기(비교 전) 상태

        private readonly Dictionary<string, object> _pendingLeft = new Dictionary<string, object>();
        private readonly Dictionary<string, object> _pendingRight = new Dictionary<string, object>();

        // 파일별(한쪽만) 저장 후 재비교 시, 저장하지 않은 방향의 병합 대기를 유지하기 위한 스냅샷.
        private Dictionary<string, object> _pendingRestoreLeft;
        private Dictionary<string, object> _pendingRestoreRight;

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

            // ===== 주 툴바(자주 쓰는 동작 — 크게) =====
            ToolStrip primary = new ToolStrip();
            _tool = primary;
            primary.GripStyle = ToolStripGripStyle.Hidden;
            primary.Dock = DockStyle.Top;
            primary.Font = new Font("Segoe UI", 12F, FontStyle.Regular);
            primary.ImageScalingSize = new Size(1, 1);
            primary.Padding = new Padding(3, 2, 3, 2);
            primary.Renderer = new ToolStripProfessionalRenderer();

            primary.Items.Add(BigBtn("좌측 열기", "비교 기준(왼쪽) 파일 열기 — Ctrl+O", delegate { OpenFile(true); }, false));
            primary.Items.Add(BigBtn("우측 열기", "비교 대상(오른쪽) 파일 열기 — Ctrl+Shift+O", delegate { OpenFile(false); }, false));
            primary.Items.Add(BigBtn("비교", "두 파일을 비교 (가장 중요) — F5", delegate { StartCompareExcel(); }, true));
            primary.Items.Add(new ToolStripSeparator());
            // 차이 이동은 화살표 없이 텍스트로.
            _btnPrev = BigBtn("이전차이", "이전 차이로 이동 (F8)", delegate { NavigateDiff(-1); }, false);
            _btnNext = BigBtn("다음차이", "다음 차이로 이동 (F7)", delegate { NavigateDiff(1); }, false);
            primary.Items.Add(_btnPrev);
            primary.Items.Add(_btnNext);
            primary.Items.Add(new ToolStripSeparator());
            // 병합: 부등호가 값이 가는 '방향'(목적지)을 가리킴.  '왼쪽 > 오른쪽' = 왼쪽값을 오른쪽으로.
            _btnMergeLR = BigBtn("왼쪽 > 오른쪽", "선택한 셀: 왼쪽(좌) 값을 오른쪽(우)에 적용(병합 대기) — Alt+→", delegate { MergeSelected(true); }, false);
            _btnMergeRL = BigBtn("왼쪽 < 오른쪽", "선택한 셀: 오른쪽(우) 값을 왼쪽(좌)에 적용(병합 대기) — Alt+←", delegate { MergeSelected(false); }, false);
            _btnSave = BigBtn("저장", "양쪽(좌·우) 병합 결과를 각 원본 파일에 모두 저장 (Ctrl+S). 한쪽만 저장하려면 경로 옆 [이 파일에 저장]을 쓰세요.", delegate { SaveMerges(); }, true);
            primary.Items.Add(_btnMergeLR);
            primary.Items.Add(_btnMergeRL);
            primary.Items.Add(_btnSave);

            // ===== 보조 툴바(옵션 — 작게) =====
            ToolStrip secondary = new ToolStrip();
            _tool2 = secondary;
            secondary.GripStyle = ToolStripGripStyle.Hidden;
            secondary.Dock = DockStyle.Top;
            secondary.Font = new Font("Segoe UI", 8.5F, FontStyle.Regular);

            // 위험 동작 강조색(연주황) — 전체병합/병합취소 구분용.
            Color dangerTint = Color.FromArgb(255, 224, 178);

            // ── 그룹: 정렬 ──
            secondary.Items.Add(GroupCaption("정렬"));
            _cboAlign = new ToolStripComboBox();
            _cboAlign.DropDownStyle = ComboBoxStyle.DropDownList;
            _cboAlign.Items.AddRange(new object[] { "좌표(빠름)", "자동정렬(LCS)", "키 컬럼" });
            _cboAlign.SelectedIndex = (int)_alignMode;
            _cboAlign.ToolTipText = "행 정렬 방식: 좌표/자동정렬(행 삽입·삭제 인지)/키 컬럼";
            _cboAlign.SelectedIndexChanged += delegate { OnAlignModeChanged(); };
            secondary.Items.Add(_cboAlign);
            secondary.Items.Add(new ToolStripLabel("키열:"));
            _txtKeyCol = new ToolStripTextBox();
            _txtKeyCol.Width = 40;
            _txtKeyCol.ToolTipText = "키 컬럼 문자(예: A). '키 컬럼' 정렬에서 사용.";
            _txtKeyCol.LostFocus += delegate { OnKeyColChanged(); };
            secondary.Items.Add(_txtKeyCol);
            secondary.Items.Add(new ToolStripSeparator());

            // ── 그룹: 보기 ──
            secondary.Items.Add(GroupCaption("보기"));
            _btnChangesOnly = new ToolStripButton("변경만 보기");
            _btnChangesOnly.CheckOnClick = true;
            _btnChangesOnly.ToolTipText = "변경/추가/삭제 행만 표시(동일 행 접기)";
            _btnChangesOnly.CheckedChanged += delegate { _changesOnly = _btnChangesOnly.Checked; RefreshGridRows(); };
            secondary.Items.Add(_btnChangesOnly);
            _btnChangeList = new ToolStripButton("변경목록 ▤");
            _btnChangeList.CheckOnClick = true;
            _btnChangeList.Checked = true;
            _btnChangeList.ToolTipText = "우측 변경목록 패널 표시/숨기기 (클릭 시 해당 셀로 이동)";
            _btnChangeList.CheckedChanged += delegate { if (_changePanel != null) _changePanel.Visible = _btnChangeList.Checked; };
            secondary.Items.Add(_btnChangeList);
            _btnFontUp = SmallBtn("글자+", "그리드 글자 크게 (Ctrl+마우스휠 위)", delegate { ApplyGridFont(_gridFontSize + 1f); });
            _btnFontDown = SmallBtn("글자−", "그리드 글자 작게 (Ctrl+마우스휠 아래)", delegate { ApplyGridFont(_gridFontSize - 1f); });
            secondary.Items.Add(_btnFontUp);
            secondary.Items.Add(_btnFontDown);
            _btnAutoFit = SmallBtn("열 자동맞춤", "표시 중인 셀 기준으로 열 너비 자동 조정 (열 머리글 더블클릭=해당 열만)", delegate { AutoFitColumns(); });
            secondary.Items.Add(_btnAutoFit);
            secondary.Items.Add(new ToolStripSeparator());

            // ── 그룹: 병합(위험 동작 — 연주황 강조) ──
            secondary.Items.Add(GroupCaption("병합"));
            _btnMergeAllLR = SmallBtn("모두 왼쪽>오른쪽", "모든 차이 셀에 대해 왼쪽 값을 오른쪽에 한꺼번에 적용", delegate { MergeAll(true); });
            _btnMergeAllRL = SmallBtn("모두 왼쪽<오른쪽", "모든 차이 셀에 대해 오른쪽 값을 왼쪽에 한꺼번에 적용", delegate { MergeAll(false); });
            _btnCancelMerge = SmallBtn("병합 취소", "병합 대기 중인 변경을 모두 취소", delegate { CancelAllMerges(); });
            _btnMergeAllLR.BackColor = dangerTint;
            _btnMergeAllRL.BackColor = dangerTint;
            _btnCancelMerge.BackColor = dangerTint;
            secondary.Items.Add(_btnMergeAllLR);
            secondary.Items.Add(_btnMergeAllRL);
            secondary.Items.Add(_btnCancelMerge);
            secondary.Items.Add(new ToolStripSeparator());

            // ── 그룹: 도구 ──
            secondary.Items.Add(GroupCaption("도구"));
            secondary.Items.Add(SmallBtn("N-way 취합…", "여러 버전(3개 이상) 취합 비교 — 사용법 안내 포함", delegate { OpenNWayDialog(); }));
            secondary.Items.Add(SmallBtn("CSV폴더비교(테스트)", "Excel 없이 CSV 폴더 2개 비교", delegate { StartCompareCsv(); }));
            secondary.Items.Add(SmallBtn("로그 열기", "로그 파일 위치 보기", delegate { ShowLogPath(); }));
            secondary.Items.Add(new ToolStripSeparator());

            // ── 그룹: 최근 비교 ──
            _btnRecent = new ToolStripDropDownButton("최근 비교 ▾");
            _btnRecent.ToolTipText = "최근 비교한 좌/우 파일 쌍 (클릭 시 불러와 비교)";
            _btnRecent.DropDownOpening += delegate { BuildRecentMenu(); };
            secondary.Items.Add(_btnRecent);

            // ----- 시트 탭
            _tabs = new TabControl();
            _tabs.Dock = DockStyle.Top;
            _tabs.Height = 26;
            _tabs.SelectedIndexChanged += delegate { OnSheetChanged(); };

            // ----- 경로 라벨 + 파일별 저장 버튼(U2: 어느 파일이 저장되는지 직관화)
            _pathTip = new ToolTip();
            Panel pathPanel = new Panel();
            pathPanel.Dock = DockStyle.Top;
            pathPanel.Height = 26;

            // 좌측 절반(경로 라벨 + [이 파일에 저장])
            Panel leftHalf = new Panel();
            leftHalf.Dock = DockStyle.Left; leftHalf.Width = 640;
            _lblLeftPath = new Label();
            _lblLeftPath.Text = "(좌측 파일 없음)";
            _lblLeftPath.Dock = DockStyle.Fill;
            _lblLeftPath.TextAlign = ContentAlignment.MiddleLeft; _lblLeftPath.AutoEllipsis = true;
            _btnSaveLeftFile = new Button();
            _btnSaveLeftFile.Text = "이 파일에 저장";
            _btnSaveLeftFile.Dock = DockStyle.Right; _btnSaveLeftFile.Width = 150;
            _btnSaveLeftFile.Enabled = false;
            _btnSaveLeftFile.Click += delegate { SaveMergesSide(true); };
            leftHalf.Controls.Add(_lblLeftPath);
            leftHalf.Controls.Add(_btnSaveLeftFile);

            // 우측 절반(경로 라벨 + [이 파일에 저장])
            Panel rightHalf = new Panel();
            rightHalf.Dock = DockStyle.Fill;
            _lblRightPath = new Label();
            _lblRightPath.Text = "(우측 파일 없음)";
            _lblRightPath.Dock = DockStyle.Fill;
            _lblRightPath.TextAlign = ContentAlignment.MiddleLeft; _lblRightPath.AutoEllipsis = true;
            _btnSaveRightFile = new Button();
            _btnSaveRightFile.Text = "이 파일에 저장";
            _btnSaveRightFile.Dock = DockStyle.Right; _btnSaveRightFile.Width = 150;
            _btnSaveRightFile.Enabled = false;
            _btnSaveRightFile.Click += delegate { SaveMergesSide(false); };
            rightHalf.Controls.Add(_lblRightPath);
            rightHalf.Controls.Add(_btnSaveRightFile);

            pathPanel.Controls.Add(rightHalf);
            pathPanel.Controls.Add(leftHalf);

            // ----- 좌우 그리드 + 마커바
            SplitContainer split = new SplitContainer();
            split.Dock = DockStyle.Fill;
            split.SplitterWidth = 4;
            _gridLeft = MakeGrid();
            _gridRight = MakeGrid();
            split.Panel1.Controls.Add(_gridLeft);

            _marker = new MarkerBar();
            _marker.Dock = DockStyle.Right;
            _marker.Width = 20; // 더 잘 보이게
            _marker.OnSeek = delegate(float pos) { SeekTo(pos); };
            split.Panel2.Controls.Add(_gridRight);
            split.Panel2.Controls.Add(_marker);

            Load += delegate { try { split.SplitterDistance = split.Width / 2; } catch { } };

            // ----- 변경 목록 패널(우측): 클릭하면 해당 셀로 점프
            _changePanel = BuildChangePanel();

            // ----- 셀 상세 패널
            Panel detailPanel = new Panel();
            detailPanel.Dock = DockStyle.Bottom;
            detailPanel.Height = 84;
            _detail = new RichTextBox();
            _detail.Dock = DockStyle.Fill;
            _detail.Font = new Font(FontFamily.GenericMonospace, 9f);
            _detail.ReadOnly = true;              // 값 선택/복사 가능(U3), 편집 불가
            _detail.BorderStyle = BorderStyle.None;
            _detail.BackColor = SystemColors.Window;
            _detail.WordWrap = false;
            _detail.ScrollBars = RichTextBoxScrollBars.Vertical;
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
            // 색상 범례를 상단에서 우측 하단(상태바 오른쪽 끝)으로 이동 — 상단 공간 회수(U4-legend).
            // _lblSummary.Spring=true 라 이후 아이템들은 자동으로 우측 정렬됨.
            AddStatusLegend(status, ColChanged, "변경");
            AddStatusLegend(status, ColAdded, "추가");
            AddStatusLegend(status, ColDeleted, "삭제");
            AddStatusLegend(status, ColMerged, "병합됨");
            AddStatusLegend(status, ColGap, "빈칸");

            // 도킹 z-order: Fill 먼저, 그 다음 Right(중앙밴드), 이어서 Top(안쪽→바깥쪽), 끝에 Bottom.
            Controls.Add(split);        // Fill
            Controls.Add(_changePanel); // Right (그리드 오른쪽, 변경목록)
            Controls.Add(pathPanel);    // Top
            Controls.Add(_tabs);        // Top
            Controls.Add(secondary);    // Top (보조, 주툴바 아래)
            Controls.Add(primary);      // Top (최상단, 주툴바)
            Controls.Add(detailPanel);  // Bottom
            Controls.Add(status);       // Bottom(최하단)

            // 그리드 셀 우클릭 컨텍스트 메뉴(U8)
            _cellMenu = new ContextMenuStrip();
            _cellMenuCancel = new ToolStripMenuItem("이 셀 병합 취소");
            _cellMenuCancel.Click += delegate { CancelOneCell(); };
            _cellMenu.Items.Add(_cellMenuCancel);

            UpdateKeyColEnabled();
            UpdateButtonStates();
            RestoreWindow();

            // 그리드 글자 크기 복원(U5). 저장값 없으면 기본 9.
            float savedFont = _settings.GridFontSize;
            ApplyGridFont(savedFont >= 6f ? savedFont : 9f);

            FormClosing += delegate { SaveWindow(); };
        }

        // ===== 변경 목록 패널 =====
        private Panel BuildChangePanel()
        {
            Panel p = new Panel();
            p.Dock = DockStyle.Right;
            p.Width = 340;

            _changeHeader = new Label();
            _changeHeader.Dock = DockStyle.Top;
            _changeHeader.Height = 22;
            _changeHeader.TextAlign = ContentAlignment.MiddleLeft;
            _changeHeader.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            _changeHeader.Text = "변경 목록 (비교 후 표시)";

            // 검색/필터 줄 (U4)
            Panel filter = new Panel();
            filter.Dock = DockStyle.Top;
            filter.Height = 26;
            _txtSearch = new TextBox();
            _txtSearch.Dock = DockStyle.Fill;
            _txtSearch.TextChanged += delegate { PopulateChangeList(); };
            _cboKind = new ComboBox();
            _cboKind.Dock = DockStyle.Right;
            _cboKind.Width = 80;
            _cboKind.DropDownStyle = ComboBoxStyle.DropDownList;
            _cboKind.Items.AddRange(new object[] { "전체", "변경", "추가", "삭제" });
            _cboKind.SelectedIndex = 0;
            _cboKind.SelectedIndexChanged += delegate { PopulateChangeList(); };
            filter.Controls.Add(_txtSearch);
            filter.Controls.Add(_cboKind);

            _changeList = new ListView();
            _changeList.Dock = DockStyle.Fill;
            _changeList.View = View.Details;
            _changeList.FullRowSelect = true;
            _changeList.MultiSelect = false;
            _changeList.HideSelection = false;
            _changeList.Columns.Add("시트", 70);
            _changeList.Columns.Add("위치", 55);
            _changeList.Columns.Add("종류", 45);
            _changeList.Columns.Add("좌 → 우", 150);
            _changeList.ItemActivate += delegate { JumpToSelectedChange(); };
            _changeList.Click += delegate { JumpToSelectedChange(); };

            p.Controls.Add(_changeList);
            p.Controls.Add(filter);
            p.Controls.Add(_changeHeader);
            return p;
        }

        /// <summary>변경목록 필터(검색어/종류)에 맞으면 true.</summary>
        private bool ChangeFilterMatch(string sheet, string pos, CellStatus st, string leftText, string rightText)
        {
            if (_cboKind != null)
            {
                int k = _cboKind.SelectedIndex;   // 0=전체 1=변경 2=추가 3=삭제
                if (k == 1 && st != CellStatus.Changed) return false;
                if (k == 2 && st != CellStatus.Added) return false;
                if (k == 3 && st != CellStatus.Deleted) return false;
            }
            if (_txtSearch != null)
            {
                string q = _txtSearch.Text != null ? _txtSearch.Text.Trim() : "";
                if (q.Length > 0)
                {
                    string hay = (sheet + " " + pos + " " + leftText + " " + rightText);
                    if (hay.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0) return false;
                }
            }
            return true;
        }

        private void PopulateChangeList()
        {
            if (_changeList == null) return;
            _changeList.BeginUpdate();
            _changeList.Items.Clear();
            int total = 0;
            if (_diff != null && !_isPreview)
            {
                const int cap = 5000;
                bool capped = false;
                foreach (SheetDiff sd in _diff.Sheets)
                {
                    if (capped) break;
                    foreach (DiffNav nv in sd.Nav)
                    {
                        if (total >= cap) { capped = true; break; }
                        DiffRow dr = sd.Rows[nv.RowIndex];
                        CellStatus st = dr.StatusAt(nv.Col);
                        object lv = dr.LeftRow >= 0 && sd.Left != null ? sd.Left.GetValueAbs(dr.LeftRow, nv.Col) : null;
                        object rv = dr.RightRow >= 0 && sd.Right != null ? sd.Right.GetValueAbs(dr.RightRow, nv.Col) : null;
                        int rowNo = dr.LeftRow >= 0 ? dr.LeftRow : dr.RightRow;

                        string pos = ColLetter(nv.Col) + rowNo;
                        string lvs = ShortCell(lv), rvs = ShortCell(rv);
                        if (!ChangeFilterMatch(sd.Name, pos, st, lvs, rvs)) continue;

                        ListViewItem it = new ListViewItem(sd.Name);
                        it.SubItems.Add(pos);
                        it.SubItems.Add(KindLabel(st));
                        it.SubItems.Add(lvs + " → " + rvs);
                        it.Tag = new object[] { sd, nv };
                        switch (st)
                        {
                            case CellStatus.Changed: it.BackColor = ColChanged; break;
                            case CellStatus.Added: it.BackColor = ColAdded; break;
                            case CellStatus.Deleted: it.BackColor = ColDeleted; break;
                        }
                        _changeList.Items.Add(it);
                        total++;
                    }
                }
                _changeHeader.Text = capped ? ("변경 목록 (" + total + "+ , 상한)") : ("변경 목록 (" + total + "건) — 클릭 시 이동");
            }
            else
            {
                _changeHeader.Text = "변경 목록 (비교 후 표시)";
            }
            _changeList.EndUpdate();
        }

        private static string ShortCell(object v)
        {
            string s = DiffEngine.ToText(v);
            if (s.Length == 0) return "(빈칸)";
            return s.Length > 24 ? s.Substring(0, 24) + "…" : s;
        }

        private void JumpToSelectedChange()
        {
            if (_changeList.SelectedItems.Count == 0) return;
            object[] tag = _changeList.SelectedItems[0].Tag as object[];
            if (tag == null) return;
            SheetDiff sd = (SheetDiff)tag[0];
            DiffNav nv = (DiffNav)tag[1];

            // 해당 시트 탭 선택.
            for (int i = 0; i < _tabs.TabPages.Count; i++)
            {
                if (_tabs.TabPages[i].Tag == sd)
                {
                    if (_tabs.SelectedIndex != i) _tabs.SelectedIndex = i; // OnSheetChanged 호출됨
                    break;
                }
            }
            // 셀 선택/스크롤.
            SelectCell(nv);
        }

        private ToolStripButton BigBtn(string text, string tip, EventHandler h, bool emphasize)
        {
            ToolStripButton b = new ToolStripButton(text);
            b.DisplayStyle = ToolStripItemDisplayStyle.Text;
            b.ToolTipText = tip;
            b.Padding = new Padding(12, 8, 12, 8);
            b.Margin = new Padding(2, 1, 2, 1);
            if (emphasize) b.Font = new Font("Segoe UI", 12F, FontStyle.Bold);
            if (h != null) b.Click += h;
            return b;
        }

        /// <summary>보조 툴바 기능 그룹 앞에 붙는 회색 캡션(정렬/보기/병합/도구).</summary>
        private static ToolStripLabel GroupCaption(string text)
        {
            ToolStripLabel l = new ToolStripLabel(text);
            l.ForeColor = Color.Gray;
            l.Font = new Font("Segoe UI", 8.5F, FontStyle.Bold);
            l.Margin = new Padding(6, 1, 2, 2);
            return l;
        }

        private ToolStripButton SmallBtn(string text, string tip, EventHandler h)
        {
            ToolStripButton b = new ToolStripButton(text);
            b.DisplayStyle = ToolStripItemDisplayStyle.Text;
            b.ToolTipText = tip;
            if (h != null) b.Click += h;
            return b;
        }

        /// <summary>상태에 따라 버튼 활성/비활성 (편의성).</summary>
        private void UpdateButtonStates()
        {
            // 미리보기 상태에서는 네비/병합/변경만 비활성(진짜 비교 후에만 의미 있음).
            bool realDiff = _diff != null && _curSheet != null && !_isPreview;
            bool canMerge = realDiff && !IsReadOnlySession();   // CSV/Word 는 저장 불가 → 병합도 비활성
            bool hasPending = (_pendingLeft.Count + _pendingRight.Count) > 0;
            if (_btnPrev != null) _btnPrev.Enabled = realDiff;
            if (_btnNext != null) _btnNext.Enabled = realDiff;
            if (_btnMergeLR != null) _btnMergeLR.Enabled = canMerge;
            if (_btnMergeRL != null) _btnMergeRL.Enabled = canMerge;
            if (_btnMergeAllLR != null) _btnMergeAllLR.Enabled = canMerge;
            if (_btnMergeAllRL != null) _btnMergeAllRL.Enabled = canMerge;
            if (_btnChangesOnly != null) _btnChangesOnly.Enabled = realDiff;
            if (_btnSave != null) _btnSave.Enabled = hasPending;
            if (_btnCancelMerge != null) _btnCancelMerge.Enabled = hasPending;

            // 파일별 저장 버튼(U2): 해당 방향 대기 있을 때만 활성 + 대기 건수 표시 + 툴팁에 전체 경로.
            bool canSave = !IsReadOnlySession();
            if (_btnSaveLeftFile != null)
            {
                int nl = _pendingLeft.Count;
                _btnSaveLeftFile.Enabled = canSave && nl > 0;
                _btnSaveLeftFile.Text = nl > 0 ? ("이 파일에 저장 (" + nl + ")") : "이 파일에 저장";
                if (_pathTip != null)
                    _pathTip.SetToolTip(_btnSaveLeftFile,
                        "좌측 파일에 저장: " + (string.IsNullOrEmpty(_leftPath) ? "(파일 없음)" : _leftPath));
            }
            if (_btnSaveRightFile != null)
            {
                int nr = _pendingRight.Count;
                _btnSaveRightFile.Enabled = canSave && nr > 0;
                _btnSaveRightFile.Text = nr > 0 ? ("이 파일에 저장 (" + nr + ")") : "이 파일에 저장";
                if (_pathTip != null)
                    _pathTip.SetToolTip(_btnSaveRightFile,
                        "우측 파일에 저장: " + (string.IsNullOrEmpty(_rightPath) ? "(파일 없음)" : _rightPath));
            }
        }

        private bool IsReadOnlySession()
        {
            return IsCsvSession() || IsWordPath(_leftPath) || IsWordPath(_rightPath);
        }

        private void ShowLogPath()
        {
            string p = Logger.LogPath ?? "(로그 파일 없음)";
            Logger.Info("사용자: 로그 위치 확인");
            MessageBox.Show(this, "로그 파일 위치:\r\n" + p +
                "\r\n\r\n오류 발생 시 이 파일을 확인/전달하세요.",
                "로그", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void RestoreWindow()
        {
            try
            {
                int w = IntSetting("WinW", 0), h = IntSetting("WinH", 0);
                if (w > 300 && h > 200) { Width = w; Height = h; }
                if (_settings.GetBool("WinMax", false)) WindowState = FormWindowState.Maximized;
            }
            catch { }
        }

        private void SaveWindow()
        {
            try
            {
                _settings.SetBool("WinMax", WindowState == FormWindowState.Maximized);
                if (WindowState == FormWindowState.Normal)
                {
                    _settings.Set("WinW", Width.ToString());
                    _settings.Set("WinH", Height.ToString());
                }
                _settings.Save();
                Logger.Info("창 상태 저장, 앱 종료 진행");
            }
            catch { }
        }

        /// <summary>상태바 오른쪽 끝에 색상 스와치 + 라벨 한 쌍을 범례로 추가.</summary>
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

        /// <summary>
        /// 더블버퍼링을 켠 DataGridView. 기본 DataGridView 는 DoubleBuffered=false 라
        /// 스크롤/리페인트마다 깜빡임·지연이 크다 → 화면 조작 속도의 가장 큰 병목.
        /// </summary>
        private sealed class BufferedGrid : DataGridView
        {
            public BufferedGrid()
            {
                DoubleBuffered = true;
                SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);
            }
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
            g.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText; // Ctrl+C 복사(U3)
            g.CellValueNeeded += Grid_CellValueNeeded;
            g.CellFormatting += Grid_CellFormatting;
            g.Scroll += Grid_Scroll;
            g.SelectionChanged += Grid_SelectionChanged;
            g.RowPostPaint += Grid_RowPostPaint;
            g.MouseWheel += Grid_MouseWheel;                              // Ctrl+휠 글자크기(U5)
            g.CellMouseDown += Grid_CellMouseDown;                        // 우클릭 병합취소(U8)
            g.ColumnHeaderMouseDoubleClick += Grid_ColHeaderDoubleClick;  // 열 자동맞춤(U9)
            return g;
        }

        // ================================================================ 파일 열기/비교
        private void OpenFile(bool left)
        {
            using (OpenFileDialog dlg = new OpenFileDialog())
            {
                dlg.Filter = "Excel/Word (*.xlsx;*.xlsm;*.xls;*.docx;*.doc;*.docm)|*.xlsx;*.xlsm;*.xls;*.docx;*.doc;*.docm"
                           + "|Excel (*.xlsx;*.xlsm;*.xls)|*.xlsx;*.xlsm;*.xls"
                           + "|Word (*.docx;*.doc;*.docm)|*.docx;*.doc;*.docm"
                           + "|모든 파일 (*.*)|*.*";
                string last = _settings.LastFolder;
                if (!string.IsNullOrEmpty(last) && Directory.Exists(last)) dlg.InitialDirectory = last;
                if (dlg.ShowDialog(this) != DialogResult.OK) return;
                if (left) { _leftPath = dlg.FileName; _lblLeftPath.Text = dlg.FileName; }
                else { _rightPath = dlg.FileName; _lblRightPath.Text = dlg.FileName; }
                Logger.Info("파일 선택(" + (left ? "좌" : "우") + "): " + dlg.FileName);
                _settings.LastFolder = Path.GetDirectoryName(dlg.FileName);
                _settings.Save();
                LoadOneAsync(dlg.FileName, left);   // 열자마자 내용 미리보기
            }
        }

        /// <summary>파일 1개를 로드해 캐시하고, 미리보기를 그린다(비교 전에도 내용 표시).</summary>
        private void LoadOneAsync(string path, bool left)
        {
            if (_busy) return;
            SetBusy(true);
            _progress.Visible = true;
            _lblSummary.Text = "로드 중… (" + Path.GetFileName(path) + ")";
            StaTask.Run<WorkbookData>(
                delegate { return LoadWorkbookSafe(path); },
                delegate(WorkbookData wb, Exception err)
                {
                    BeginInvoke((MethodInvoker)delegate { OnOneLoaded(wb, err, left); });
                });
        }

        private void OnOneLoaded(WorkbookData wb, Exception err, bool left)
        {
            _progress.Visible = false;
            SetBusy(false);
            if (err != null)
            {
                Logger.Error("파일 로드 실패(미리보기)", err);
                _lblSummary.Text = "로드 오류: " + err.Message;
                MessageBox.Show(this, DescribeError(err), "로드 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (left) _leftWb = wb; else _rightWb = wb;
            _diff = null;               // 이전 비교 결과 무효화
            _pendingLeft.Clear();
            _pendingRight.Clear();
            ShowPreview();
        }

        /// <summary>로드된 워크북(들)을 색상 없이 원본 그대로 좌/우 그리드에 표시.</summary>
        private void ShowPreview()
        {
            if (_leftWb == null && _rightWb == null) return;
            _isPreview = true;
            _diff = BuildPreviewResult(_leftWb, _rightWb);
            PopulateTabs();
            PopulateChangeList();
            string ls = _leftWb != null ? Path.GetFileName(_leftWb.FilePath) : "(없음)";
            string rs = _rightWb != null ? Path.GetFileName(_rightWb.FilePath) : "(없음)";
            _lblSummary.Text = "미리보기 — 좌:" + ls + "  우:" + rs + "   [비교]를 누르면 차이를 표시합니다.";
            Logger.Info("미리보기 표시 (좌=" + ls + ", 우=" + rs + ")");
            UpdateButtonStates();
        }

        /// <summary>비교 전 미리보기용 DiffResult(좌표 정렬, 변경표시 없음)을 만든다.</summary>
        private static DiffResult BuildPreviewResult(WorkbookData left, WorkbookData right)
        {
            DiffResult res = new DiffResult();
            res.LeftPath = left != null ? left.FilePath : "";
            res.RightPath = right != null ? right.FilePath : "";

            List<string> names = new List<string>();
            HashSet<string> seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            if (left != null) foreach (SheetData s in left.Sheets) if (seen.Add(s.Name)) names.Add(s.Name);
            if (right != null) foreach (SheetData s in right.Sheets) if (seen.Add(s.Name)) names.Add(s.Name);

            foreach (string name in names)
            {
                SheetData l = left != null ? left.FindSheet(name) : null;
                SheetData r = right != null ? right.FindSheet(name) : null;
                SheetDiff sd = new SheetDiff();
                sd.Name = name; sd.Left = l; sd.Right = r; sd.Status = SheetStatus.Same;

                int minCol = int.MaxValue, maxCol = 0, minRow = int.MaxValue, maxRow = 0;
                if (l != null && l.ColCount > 0) { minCol = Math.Min(minCol, l.FirstCol); maxCol = Math.Max(maxCol, l.LastCol); minRow = Math.Min(minRow, l.FirstRow); maxRow = Math.Max(maxRow, l.LastRow); }
                if (r != null && r.ColCount > 0) { minCol = Math.Min(minCol, r.FirstCol); maxCol = Math.Max(maxCol, r.LastCol); minRow = Math.Min(minRow, r.FirstRow); maxRow = Math.Max(maxRow, r.LastRow); }
                if (maxCol < minCol) { minCol = 1; maxCol = 0; minRow = 1; maxRow = 0; }
                sd.MinCol = minCol; sd.MaxCol = maxCol;

                for (int absRow = minRow; absRow <= maxRow; absRow++)
                {
                    int lr = (l != null && absRow >= l.FirstRow && absRow <= l.LastRow) ? absRow : -1;
                    int rr = (r != null && absRow >= r.FirstRow && absRow <= r.LastRow) ? absRow : -1;
                    DiffRow dr = new DiffRow(lr, rr);
                    dr.Kind = RowKind.Same;   // 미리보기 → 색상 없음
                    sd.Rows.Add(dr);
                }
                res.Sheets.Add(sd);
            }
            return res;
        }

        private void StartCompareExcel()
        {
            if (string.IsNullOrEmpty(_leftPath) || string.IsNullOrEmpty(_rightPath))
            {
                Info("좌측과 우측 파일을 모두 지정하세요."); return;
            }
            // 이미 열어둔(캐시된) 워크북이 있으면 재로드 없이 바로 비교 → 빠르고 COM 재열기 없음.
            if (_leftWb != null && _rightWb != null)
            {
                Logger.Info("비교(캐시 사용): 좌=" + _leftPath + " 우=" + _rightPath);
                RunDiff();
                return;
            }
            LoadAndCompare(false);
        }

        private void StartCompareCsv()
        {
            DialogResult ok = MessageBox.Show(this,
                "CSV 폴더 비교 (개발/테스트용 — Excel 없이 동작)\r\n\r\n"
                + "· 비교할 CSV 파일들을 담은 폴더를 좌/우 각각 고릅니다.\r\n"
                + "· 폴더 안의 .csv 파일 1개 = 시트 1개로 취급합니다.\r\n"
                + "   예) old\\Sheet1.csv, old\\Sheet2.csv  ↔  new\\Sheet1.csv, new\\Sheet2.csv\r\n"
                + "· 실제 Excel/DRM 파일 비교는 [좌측 열기]/[우측 열기]를 쓰세요.\r\n\r\n"
                + "계속하시겠어요?",
                "CSV 폴더 비교", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
            if (ok != DialogResult.OK) return;

            string l = PickFolder("좌측 CSV 폴더 선택 (각 .csv = 시트)");
            if (l == null) return;
            string r = PickFolder("우측 CSV 폴더 선택");
            if (r == null) return;
            _leftPath = l; _rightPath = r;
            _lblLeftPath.Text = l; _lblRightPath.Text = r;
            _leftWb = null; _rightWb = null; // 폴더 비교는 캐시 무효화 후 로드
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
            Logger.Info("비교 시작 (" + (useCsv ? "CSV" : "Excel COM") + ")  좌=" + _leftPath + "  우=" + _rightPath);
            _pendingLeft.Clear();
            _pendingRight.Clear();
            _progress.Visible = true;
            _lblSummary.Text = "로드 중…";
            SetBusy(true);

            string lp = _leftPath, rp = _rightPath;
            StaTask.Run<WorkbookData[]>(
                delegate
                {
                    // 파일마다 '자기 리더'로 열고 닫는다 → 한 Excel 세션에서 두 워크북을
                    // 연달아 여닫을 때 생기던 COM 상태 문제를 근본 회피(로그의 로드 실패 대응).
                    WorkbookData l = LoadWorkbookSafe(lp);
                    WorkbookData r = LoadWorkbookSafe(rp);
                    return new WorkbookData[] { l, r };
                },
                delegate(WorkbookData[] res, Exception err)
                {
                    // STA 워커 스레드 → UI 스레드로 마샬링.
                    BeginInvoke((MethodInvoker)delegate { OnLoaded(res, err); });
                });
        }

        /// <summary>경로 1개를 확장자에 맞는 리더로 로드. 리더는 매번 새로 생성/해제.</summary>
        private static WorkbookData LoadWorkbookSafe(string path)
        {
            using (IWorkbookReader reader = MakeReader(path))
            {
                return reader.LoadWorkbook(path);
            }
        }

        private static IWorkbookReader MakeReader(string path)
        {
            if (Directory.Exists(path)) return new CsvWorkbookReader();   // CSV 폴더(테스트)
            if (IsWordPath(path)) return new WordComReader();             // Word 문서
            return new ExcelComReader();                                  // Excel(기본)
        }

        private static bool IsWordPath(string path)
        {
            if (string.IsNullOrEmpty(path)) return false;
            string ext = Path.GetExtension(path).ToLowerInvariant();
            return ext == ".doc" || ext == ".docx" || ext == ".docm";
        }

        private void OnLoaded(WorkbookData[] res, Exception err)
        {
            _progress.Visible = false;
            SetBusy(false);
            if (err != null)
            {
                Logger.Error("로드 실패", err);
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

            // B3: 키 컬럼 모드인데 키열이 비었거나 잘못되면 조용히 LCS 폴백하지 않고 안내.
            AlignMode effMode = _alignMode;
            if (_alignMode == AlignMode.KeyColumn && _keyCol < 0)
            {
                DialogResult a = MessageBox.Show(this,
                    "키열이 비어 있거나 잘못됨 — 예: A\r\n\r\n"
                    + "[예] 자동정렬(LCS)로 대신 비교합니다.\r\n"
                    + "[아니오] 비교를 중단합니다(키열을 입력 후 다시 시도).",
                    "키 컬럼 오류", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (a != DialogResult.Yes)
                {
                    _lblSummary.Text = "비교 중단 — 키열이 비어 있거나 잘못됨(예: A).";
                    Logger.Info("비교 중단: 키 컬럼 모드지만 키열 파싱 실패");
                    return;
                }
                effMode = AlignMode.Auto;   // 사용자 확인 → 자동정렬로 진행
                Logger.Info("키열 미입력 → 사용자 확인 후 자동정렬(LCS)로 진행");
            }

            _isPreview = false;
            _pendingLeft.Clear();
            _pendingRight.Clear();
            // 파일별(한쪽만) 저장 직후 재비교라면, 저장하지 않은 방향의 대기를 복원.
            if (_pendingRestoreLeft != null)
            {
                foreach (KeyValuePair<string, object> kv in _pendingRestoreLeft) _pendingLeft[kv.Key] = kv.Value;
                _pendingRestoreLeft = null;
            }
            if (_pendingRestoreRight != null)
            {
                foreach (KeyValuePair<string, object> kv in _pendingRestoreRight) _pendingRight[kv.Key] = kv.Value;
                _pendingRestoreRight = null;
            }
            try
            {
                _diff = DiffEngine.Compare(_leftWb, _rightWb, effMode, _keyCol);
            }
            catch (Exception ex)
            {
                Logger.Error("diff 계산 실패", ex);
                _lblSummary.Text = "비교 오류: " + ex.Message;
                return;
            }
            PopulateTabs();
            PopulateChangeList();
            _lblSummary.Text = _diff.Summary() + "  [" + AlignNameFor(effMode) + "]";
            Logger.Info("비교 결과: " + _diff.Summary() + " [" + AlignNameFor(effMode) + "]");
            UpdateButtonStates();
        }

        private string AlignNameFor(AlignMode mode)
        {
            if (mode == AlignMode.Coordinate) return "좌표";
            if (mode == AlignMode.KeyColumn) return "키:" + (_keyCol >= 0 ? ColLetter(_keyCol) : "?");
            return "자동정렬";
        }

        private void OnAlignModeChanged()
        {
            _alignMode = (AlignMode)_cboAlign.SelectedIndex;
            _settings.Set("AlignMode", ((int)_alignMode).ToString());
            _settings.Save();
            UpdateKeyColEnabled();
            // 키 컬럼 모드로 막 전환했는데 키열이 아직 없으면, 곧바로 경고창을 띄우지 않고
            // 안내만 한다(사용자가 키열 입력 후 [비교] 시 검증).
            if (_alignMode == AlignMode.KeyColumn && _keyCol < 0 && !_isPreview)
            {
                _lblSummary.Text = "키 컬럼 모드 — 키열(예: A)을 입력한 뒤 [비교]를 누르세요.";
                return;
            }
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
            if (_tool2 != null) _tool2.Enabled = !busy;
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
            // 단일 열(예: Word 문단 뷰)은 넓게 → 텍스트가 잘 보이게.
            int width = _colCount == 1 ? 640 : 110;
            for (int i = 0; i < _colCount; i++)
            {
                DataGridViewTextBoxColumn col = new DataGridViewTextBoxColumn();
                col.HeaderText = _colCount == 1 ? "본문" : ColLetter(_colStart + i);
                col.Width = width;
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
            // 속도: 병합 대기가 하나도 없으면 셀마다 문자열 키를 만들지 않는다(GC 압박 제거).
            Dictionary<string, object> pend = isLeft ? _pendingLeft : _pendingRight;
            if (pend.Count > 0)
            {
                object pv;
                if (pend.TryGetValue(CellKey(_curSheet.Name, absRow, absCol), out pv)) return pv;
            }
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
                // 속도: 병합 대기가 있을 때만 셀별 문자열 키 생성·조회(대부분은 대기 0 → 건너뜀).
                Dictionary<string, object> pend = isLeft ? _pendingLeft : _pendingRight;
                if (pend.Count > 0 && pend.ContainsKey(CellKey(_curSheet.Name, absRow, absCol)))
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

        // 반대편 그리드에 미러링할 때 재진입(SelectionChanged 연쇄)을 막는 가드.
        private const int MirrorCellCap = 4000;

        private void Grid_SelectionChanged(object sender, EventArgs e)
        {
            if (_syncing) return;
            _syncing = true;
            try
            {
                DataGridView src = (DataGridView)sender;
                DataGridView dst = (src == _gridLeft) ? _gridRight : _gridLeft;
                MirrorSelection(src, dst);
                if (src.CurrentCell != null)
                {
                    int r = src.CurrentCell.RowIndex, c = src.CurrentCell.ColumnIndex;
                    // 속도: 같은 셀에 대한 중복 SelectionChanged(미러링 연쇄 등)면 상세 재계산 생략.
                    if (r != _lastDetailRow || c != _lastDetailCol)
                    {
                        _lastDetailRow = r; _lastDetailCol = c;
                        UpdateDetail(r, c);
                    }
                }
            }
            catch { }
            finally { _syncing = false; }
        }

        /// <summary>B1: 소스 그리드의 '선택 셀 전체'를 반대편 그리드에 동일하게 미러링.</summary>
        private void MirrorSelection(DataGridView src, DataGridView dst)
        {
            if (dst.RowCount == 0 || dst.ColumnCount == 0) return;

            DataGridViewSelectedCellCollection sel = src.SelectedCells;

            // 빠른 경로: 단일 셀 선택(대부분의 클릭). ClearSelection+루프 없이 현재 셀만 맞춘다.
            if (sel.Count <= 1)
            {
                if (src.CurrentCell != null)
                {
                    int cr = src.CurrentCell.RowIndex, cc = src.CurrentCell.ColumnIndex;
                    if (cr >= 0 && cr < dst.RowCount && cc >= 0 && cc < dst.ColumnCount)
                    {
                        try
                        {
                            dst.CurrentCell = dst.Rows[cr].Cells[cc];
                            if (!dst.Rows[cr].Cells[cc].Selected)
                            {
                                dst.ClearSelection();
                                dst.Rows[cr].Cells[cc].Selected = true;
                            }
                        }
                        catch { }
                    }
                }
                return;
            }

            // 다중 선택 경로: CurrentCell 을 먼저 세팅(그 뒤 다중 선택이 덮이지 않도록).
            if (src.CurrentCell != null)
            {
                int cr = src.CurrentCell.RowIndex, cc = src.CurrentCell.ColumnIndex;
                if (cr >= 0 && cr < dst.RowCount && cc >= 0 && cc < dst.ColumnCount)
                {
                    try { dst.CurrentCell = dst.Rows[cr].Cells[cc]; }
                    catch { }
                }
            }
            // 성능: 수천 셀 초과 선택 시 전체 미러링은 생략(CurrentCell 만 반영).
            // 병합은 MergeSelected 가 '더 많이 선택된' 그리드를 소스로 쓰므로 정확성은 유지됨.
            if (sel.Count > MirrorCellCap)
            {
                dst.ClearSelection();
                if (src.CurrentCell != null)
                {
                    int cr = src.CurrentCell.RowIndex, cc = src.CurrentCell.ColumnIndex;
                    if (cr >= 0 && cr < dst.RowCount && cc >= 0 && cc < dst.ColumnCount)
                        dst.Rows[cr].Cells[cc].Selected = true;
                }
                return;
            }

            dst.ClearSelection();
            for (int i = 0; i < sel.Count; i++)
            {
                DataGridViewCell cell = sel[i];
                int r = cell.RowIndex, c = cell.ColumnIndex;
                if (r >= 0 && r < dst.RowCount && c >= 0 && c < dst.ColumnCount)
                    dst.Rows[r].Cells[c].Selected = true;
            }
        }

        private void UpdateDetail(int displayRow, int col)
        {
            if (_detail == null) return;
            _detail.Clear();
            if (_curSheet == null || displayRow < 0 || displayRow >= _rowMap.Length || col < 0)
                return;

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
            string lt = Trunc(DiffEngine.ToText(lv));
            string rt = Trunc(DiffEngine.ToText(rv));

            // 헤더 줄.
            RtAppend(string.Format("[{0}]  {1}열  (좌 {2}행 / 우 {3}행)\r\n", KindLabel(st), addr, lrow, rrow), null);

            // 변경 셀만 가운데 변경 구간을 강조(U7). 공통 접두/접미 제외.
            int pre = 0, suf = 0;
            bool highlight = (st == CellStatus.Changed) && lt != rt;
            if (highlight) CommonAffix(lt, rt, out pre, out suf);

            RtAppend("좌 값: ", null);
            RtAppendSegmented(lt, highlight, pre, suf, ColDeleted);
            RtAppend("   |   좌 수식: " + Trunc(FormulaText(lf)) + "\r\n", null);

            RtAppend("우 값: ", null);
            RtAppendSegmented(rt, highlight, pre, suf, ColAdded);
            RtAppend("   |   우 수식: " + Trunc(FormulaText(rf)), null);

            _detail.Select(0, 0);
            _detail.SelectionLength = 0;
        }

        /// <summary>RichTextBox 에 배경색(back=null 이면 기본)으로 텍스트 추가.</summary>
        private void RtAppend(string text, Color? back)
        {
            if (string.IsNullOrEmpty(text)) return;
            int start = _detail.TextLength;
            _detail.AppendText(text);
            _detail.Select(start, text.Length);
            _detail.SelectionBackColor = back.HasValue ? back.Value : _detail.BackColor;
            _detail.SelectionColor = _detail.ForeColor;
            _detail.SelectionLength = 0;
        }

        /// <summary>공통 접두(pre)/접미(suf)를 제외한 가운데 구간만 배경 강조해 추가.</summary>
        private void RtAppendSegmented(string text, bool highlight, int pre, int suf, Color hl)
        {
            if (text == null) text = "";
            if (!highlight || pre + suf >= text.Length)
            {
                RtAppend(text.Length == 0 ? "(빈칸)" : text, null);
                return;
            }
            string a = text.Substring(0, pre);
            string b = text.Substring(pre, text.Length - suf - pre);
            string c = text.Substring(text.Length - suf);
            RtAppend(a, null);
            RtAppend(b, hl);
            RtAppend(c, null);
        }

        /// <summary>두 문자열의 공통 접두/접미 길이(겹치지 않게).</summary>
        private static void CommonAffix(string a, string b, out int pre, out int suf)
        {
            int max = Math.Min(a.Length, b.Length);
            int p = 0;
            while (p < max && a[p] == b[p]) p++;
            int s = 0;
            while (s < (max - p) && a[a.Length - 1 - s] == b[b.Length - 1 - s]) s++;
            pre = p; suf = s;
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
            else if (e.KeyCode == Keys.F5) { StartCompareExcel(); e.Handled = true; }
            else if (e.Control && e.Shift && e.KeyCode == Keys.O) { OpenFile(false); e.Handled = true; e.SuppressKeyPress = true; }
            else if (e.Control && e.KeyCode == Keys.S) { SaveMerges(); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.O) { OpenFile(true); e.Handled = true; }
            else if (e.Control && e.KeyCode == Keys.F) { FocusChangeSearch(); e.Handled = true; e.SuppressKeyPress = true; }
            else if (e.Alt && e.KeyCode == Keys.Right) { MergeSelected(true); e.Handled = true; e.SuppressKeyPress = true; }
            else if (e.Alt && e.KeyCode == Keys.Left) { MergeSelected(false); e.Handled = true; e.SuppressKeyPress = true; }
        }

        /// <summary>Ctrl+F → 변경목록 검색창으로 포커스(U4).</summary>
        private void FocusChangeSearch()
        {
            if (_txtSearch == null) return;
            if (_changePanel != null && !_changePanel.Visible)
            {
                _changePanel.Visible = true;
                if (_btnChangeList != null) _btnChangeList.Checked = true;
            }
            _txtSearch.Focus();
            _txtSearch.SelectAll();
        }

        private void NavigateDiff(int dir)
        {
            if (_diff == null || _isPreview || _curSheet == null) return;

            // 현재 시트 안에서 이동 가능한지 먼저 시도.
            if (_curSheet.Nav.Count > 0)
            {
                int ni = _navIndex + dir;
                if (ni >= 0 && ni < _curSheet.Nav.Count)
                {
                    _navIndex = ni;
                    SelectCell(_curSheet.Nav[_navIndex]);
                    return;
                }
            }

            // 현재 시트 경계 도달 → 변경이 있는 다음/이전 시트로 자동 전환(U6).
            if (JumpToAdjacentSheetDiff(dir)) return;

            // 다른 시트에 변경이 없으면 현재 시트 안에서 순환(기존 동작).
            if (_curSheet.Nav.Count == 0) return;
            _navIndex = dir > 0 ? 0 : _curSheet.Nav.Count - 1;
            SelectCell(_curSheet.Nav[_navIndex]);
        }

        /// <summary>변경이 있는 다음/이전 시트 탭으로 전환해 그 시트의 첫/마지막 차이로 이동.</summary>
        private bool JumpToAdjacentSheetDiff(int dir)
        {
            if (_diff == null || _diff.Sheets.Count == 0) return false;
            int cur = _diff.Sheets.IndexOf(_curSheet);
            if (cur < 0) return false;
            int n = _diff.Sheets.Count;
            for (int step = 1; step <= n; step++)
            {
                int idx = cur + dir * step;
                idx = ((idx % n) + n) % n;   // 순환
                if (idx == cur) break;
                SheetDiff sd = _diff.Sheets[idx];
                if (sd.Nav.Count == 0) continue;
                SelectSheetTab(sd);          // OnSheetChanged → _curSheet=sd, _navIndex=-1
                if (_curSheet == null) return false;
                _navIndex = dir > 0 ? 0 : _curSheet.Nav.Count - 1;
                SelectCell(_curSheet.Nav[_navIndex]);
                return true;
            }
            return false;
        }

        private void SelectSheetTab(SheetDiff sd)
        {
            for (int i = 0; i < _tabs.TabPages.Count; i++)
            {
                if (_tabs.TabPages[i].Tag == sd)
                {
                    if (_tabs.SelectedIndex != i) _tabs.SelectedIndex = i; // OnSheetChanged 호출됨
                    break;
                }
            }
        }

        /// <summary>주어진 차이(Nav) 셀을 좌/우 그리드에서 선택하고 화면에 보이게 스크롤.</summary>
        private void SelectCell(DiffNav nv)
        {
            if (_curSheet == null) return;
            int displayRow;
            if (!_rowsIdxToDisplay.TryGetValue(nv.RowIndex, out displayRow))
            {
                // "변경만 보기" 상태 등으로 매핑 실패 시 필터를 끄고 재시도.
                if (_changesOnly)
                {
                    _btnChangesOnly.Checked = false;
                    _changesOnly = false;
                    RefreshGridRows();
                }
                if (!_rowsIdxToDisplay.TryGetValue(nv.RowIndex, out displayRow)) return;
            }
            int displayCol = nv.Col - _colStart;
            if (displayCol < 0 || displayCol >= _colCount) return;
            try
            {
                _gridLeft.CurrentCell = _gridLeft.Rows[displayRow].Cells[displayCol];
                _gridLeft.FirstDisplayedScrollingRowIndex = Math.Max(0, displayRow - 3);
            }
            catch { }
            // 네비 인덱스를 이 셀에 맞춰 동기화(F7/F8 이 이어지도록).
            int idx = _curSheet.Nav.FindIndex(delegate(DiffNav x) { return x.RowIndex == nv.RowIndex && x.Col == nv.Col; });
            if (idx >= 0) _navIndex = idx;

            // 전역 진행도(U6): 시트 경계를 넘는 네비게이션도 한눈에.
            int globalIdx, globalTotal, sheetNo, sheetTotal;
            GlobalNavInfo(out globalIdx, out globalTotal, out sheetNo, out sheetTotal);
            _lblSummary.Text = string.Format(
                "차이 {0}/{1} — 시트 {2}/{3}  |  이 시트 {4}/{5}  ({6}열, 좌{7}/우{8}행)",
                globalIdx, globalTotal, sheetNo, sheetTotal,
                _navIndex + 1, _curSheet.Nav.Count, ColLetter(nv.Col),
                RowLabel(true, nv.RowIndex), RowLabel(false, nv.RowIndex));
        }

        /// <summary>전체 시트에 걸친 차이 진행도(현재 차이 전역 순번/총 차이수, 현재 시트 순번/총 시트수).</summary>
        private void GlobalNavInfo(out int globalIdx, out int globalTotal, out int sheetNo, out int sheetTotal)
        {
            globalIdx = 0; globalTotal = 0; sheetNo = 1; sheetTotal = 0;
            if (_diff == null) return;
            sheetTotal = _diff.Sheets.Count;
            int before = 0;
            bool passed = false;
            for (int i = 0; i < _diff.Sheets.Count; i++)
            {
                SheetDiff sd = _diff.Sheets[i];
                if (sd == _curSheet)
                {
                    sheetNo = i + 1;
                    before = globalTotal;   // 현재 시트 앞까지의 누적 차이 수
                    passed = true;
                }
                globalTotal += sd.Nav.Count;
            }
            if (!passed) before = 0;
            globalIdx = before + _navIndex + 1;
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

            // B1: 병합 방향은 버튼이 정하고, '선택 좌표'는 더 많이 선택된 그리드에서 취한다.
            //     (미러링이 상한으로 접혔을 때도 반대편에서 드래그한 다중 선택을 살린다.)
            DataGridView selGrid = (_gridLeft.SelectedCells.Count >= _gridRight.SelectedCells.Count)
                                 ? _gridLeft : _gridRight;
            if (selGrid.SelectedCells.Count == 0 && selGrid.CurrentCell != null)
                selGrid.CurrentCell.Selected = true;

            int applied = 0, skipped = 0;
            foreach (DataGridViewCell cell in selGrid.SelectedCells)
            {
                if (cell.RowIndex < 0 || cell.RowIndex >= _rowMap.Length) continue;
                DiffRow dr = _curSheet.Rows[_rowMap[cell.RowIndex]];
                int absCol = _colStart + cell.ColumnIndex;

                int srcRow = leftToRight ? dr.LeftRow : dr.RightRow;
                int dstRow = leftToRight ? dr.RightRow : dr.LeftRow;
                if (dstRow < 0 || srcRow < 0)
                {
                    // 대상행 또는 소스행이 없음(행 추가/삭제) → 미지원.
                    // (srcRow<0 을 null 로 등록하면 반대편 신규 데이터를 지우므로 제외.)
                    skipped++;
                    continue;
                }
                object srcVal = EffectiveValue(leftToRight, dr, absCol);
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
            Logger.Info(string.Format("병합 대기 등록 {0} (적용 {1}, 미지원 {2}) 시트={3}",
                leftToRight ? "좌→우" : "우→좌", applied, skipped, _curSheet != null ? _curSheet.Name : "?"));
            UpdateButtonStates();
        }

        /// <summary>모든 시트의 모든 차이 셀을 한쪽 값으로 일괄 병합(대기 등록). leftToRight=true → 왼쪽값을 오른쪽에.</summary>
        private void MergeAll(bool leftToRight)
        {
            if (_diff == null || _isPreview) { Info("먼저 [비교]를 실행하세요."); return; }
            if (IsReadOnlySession()) { Info("CSV/Word 세션은 병합 저장을 지원하지 않습니다."); return; }

            // 적용 가능 개수 미리 집계(행추가/삭제로 대상행 없는 셀은 제외).
            int applicable = 0, skipped = 0;
            foreach (SheetDiff sd in _diff.Sheets)
            {
                foreach (DiffNav nv in sd.Nav)
                {
                    DiffRow dr = sd.Rows[nv.RowIndex];
                    int srcRow = leftToRight ? dr.LeftRow : dr.RightRow;
                    int dstRow = leftToRight ? dr.RightRow : dr.LeftRow;
                    // B2: 대상행뿐 아니라 '소스행 없음(srcRow<0)'도 제외. (null 등록 시 반대편 신규 데이터 삭제됨.)
                    if (dstRow < 0 || srcRow < 0) skipped++; else applicable++;
                }
            }
            if (applicable == 0)
            {
                Info("일괄 병합할 셀이 없습니다." + (skipped > 0 ? "\r\n(행 추가/삭제 " + skipped + "건은 미지원)" : ""));
                return;
            }

            string dir = leftToRight ? "왼쪽값 → 오른쪽" : "오른쪽값 → 왼쪽";
            DialogResult ans = MessageBox.Show(this,
                "모든 시트의 차이 " + applicable + "개 셀을 [" + dir + "] 방향으로 한꺼번에 병합 대기에 등록합니다.\r\n"
                + (skipped > 0 ? "(행 추가/삭제로 대응 불가한 " + skipped + "건은 제외)\r\n" : "")
                + "\r\n계속하시겠어요? (등록 후 [저장]을 눌러야 실제 파일에 반영됩니다.)",
                "전체 병합", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (ans != DialogResult.Yes) return;

            int applied = 0;
            foreach (SheetDiff sd in _diff.Sheets)
            {
                foreach (DiffNav nv in sd.Nav)
                {
                    DiffRow dr = sd.Rows[nv.RowIndex];
                    int srcRow = leftToRight ? dr.LeftRow : dr.RightRow;
                    int dstRow = leftToRight ? dr.RightRow : dr.LeftRow;
                    // B2: 소스행이 없으면(srcRow<0) 건너뜀 — null 로 덮으면 반대편 신규 행 데이터가 지워짐.
                    if (dstRow < 0 || srcRow < 0) continue;
                    SheetData srcSheet = leftToRight ? sd.Left : sd.Right;
                    object srcVal = srcSheet != null ? srcSheet.GetValueAbs(srcRow, nv.Col) : null;
                    string key = CellKey(sd.Name, dstRow, nv.Col);
                    if (leftToRight) _pendingRight[key] = srcVal;
                    else _pendingLeft[key] = srcVal;
                    applied++;
                }
            }
            _gridLeft.Invalidate();
            _gridRight.Invalidate();
            _lblSummary.Text = string.Format("전체 병합 대기 등록 완료 ({0}개 {1}). [저장]으로 반영하세요. 대기 좌:{2} 우:{3}",
                applied, dir, _pendingLeft.Count, _pendingRight.Count);
            Logger.Info(string.Format("전체 병합 {0}: 적용 {1}, 제외 {2}", dir, applied, skipped));
            UpdateButtonStates();
        }

        /// <summary>툴바 [저장] — 양쪽(대기 있는 방향 모두) 저장(현행 동작).</summary>
        private void SaveMerges()
        {
            DoSave(true, true);
        }

        /// <summary>경로 라벨 옆 [이 파일에 저장] — 해당 방향만 저장(U2).</summary>
        private void SaveMergesSide(bool left)
        {
            DoSave(left, !left);
        }

        /// <summary>지정한 방향(들)의 병합 대기를 원본 파일에 저장. doLeft/doRight 로 대상 선택.</summary>
        private void DoSave(bool doLeft, bool doRight)
        {
            bool saveLeft = doLeft && _pendingLeft.Count > 0;
            bool saveRight = doRight && _pendingRight.Count > 0;
            if (!saveLeft && !saveRight)
            {
                Info("저장할 병합 변경이 없습니다.\r\n\r\n먼저 셀을 선택하고 [왼쪽 > 오른쪽]/[왼쪽 < 오른쪽]으로 병합하거나, 보조 툴바의 [전체병합]을 사용한 뒤 저장하세요.");
                return;
            }
            if (IsCsvSession())
            {
                Info("CSV 테스트 세션은 저장을 지원하지 않습니다(실제 Excel 파일에서만 병합 저장).");
                return;
            }
            if (IsWordPath(_leftPath) || IsWordPath(_rightPath))
            {
                Info("Word(.docx) 문서는 현재 비교/보기 전용입니다(병합 저장 미지원).\r\n"
                   + "Word 병합이 필요하면 알려주세요 — 다음 버전에서 지원하겠습니다.");
                return;
            }

            // 어떤 파일이 어디에 저장되는지 명확히 안내.
            List<string> targets = new List<string>();
            if (saveRight) targets.Add("우측 파일 ← " + _rightPath + "   (" + _pendingRight.Count + "셀 변경)");
            if (saveLeft) targets.Add("좌측 파일 ← " + _leftPath + "   (" + _pendingLeft.Count + "셀 변경)");
            string msg = "아래 원본 파일을 덮어써 저장합니다:\r\n\r\n · "
                       + string.Join("\r\n · ", targets.ToArray())
                       + "\r\n\r\n[예] 백업 복사본을 만든 뒤 저장\r\n[아니오] 백업 없이 저장\r\n[취소] 저장 안 함"
                       + "\r\n\r\n※ 해당 원본이 Excel에서 열려 있으면 저장이 실패합니다. 먼저 닫아 주세요.";
            DialogResult dr = MessageBox.Show(this, msg, "저장 확인", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);
            if (dr == DialogResult.Cancel) return;
            bool backup = (dr == DialogResult.Yes);

            Logger.Info(string.Format("저장 시작 (좌 {0}, 우 {1}, 백업={2})",
                saveLeft ? _pendingLeft.Count : 0, saveRight ? _pendingRight.Count : 0, backup));

            bool ok = false;
            List<string> saved = new List<string>();
            try
            {
                SetBusy(true);
                if (saveRight)
                {
                    string bp = SaveOneSide(_rightPath, _pendingRight, backup);
                    saved.Add("우: " + _rightPath + (bp != null ? "\r\n     (백업: " + bp + ")" : ""));
                }
                if (saveLeft)
                {
                    string bp = SaveOneSide(_leftPath, _pendingLeft, backup);
                    saved.Add("좌: " + _leftPath + (bp != null ? "\r\n     (백업: " + bp + ")" : ""));
                }
                if (saveLeft) _pendingLeft.Clear();
                if (saveRight) _pendingRight.Clear();
                Logger.Info("저장 완료: " + string.Join(" | ", saved.ToArray()));
                ok = true;
            }
            catch (Exception ex)
            {
                Logger.Error("저장 실패", ex);
                MessageBox.Show(this,
                    "저장 실패 — 원본은 그대로 보존되었습니다.\r\n\r\n" + ex.Message +
                    "\r\n\r\n가능 원인:\r\n"
                    + " · 원본 파일이 Excel 에서 열려 있음 → 닫고 다시 시도\r\n"
                    + " · DRM 편집권한 없음(열람만 가능한 파일)\r\n"
                    + " · 다른 프로그램이 파일을 잠금",
                    "저장 실패", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                SetBusy(false);
                UpdateButtonStates();
                _gridLeft.Invalidate();
                _gridRight.Invalidate();
            }

            if (ok)
            {
                // 부분 저장(한쪽만)이면 저장하지 않은 방향의 대기를 재비교 후에도 유지(데이터 유실 방지).
                _pendingRestoreLeft = (!saveLeft && _pendingLeft.Count > 0)
                    ? new Dictionary<string, object>(_pendingLeft) : null;
                _pendingRestoreRight = (!saveRight && _pendingRight.Count > 0)
                    ? new Dictionary<string, object>(_pendingRight) : null;

                _lblSummary.Text = "저장 완료 — 최신 내용으로 다시 불러오는 중…";
                MessageBox.Show(this,
                    "저장이 완료되었습니다:\r\n\r\n" + string.Join("\r\n", saved.ToArray())
                    + "\r\n\r\n저장된 내용을 반영하기 위해 다시 비교합니다.",
                    "저장 완료", MessageBoxButtons.OK, MessageBoxIcon.Information);
                // 화면 갱신: 디스크에서 재로드 후 재비교(저장 후 화면 미갱신 문제 해결).
                LoadAndCompare(false);
            }
        }

        private bool IsCsvSession()
        {
            return Directory.Exists(_leftPath) || Directory.Exists(_rightPath);
        }

        /// <summary>한쪽 파일을 저장하고 백업 경로(없으면 null)를 반환.</summary>
        private string SaveOneSide(string path, Dictionary<string, object> pending, bool backup)
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
            string backupOut = null;
            Thread t = StaTask.Run<bool>(
                delegate
                {
                    using (MergeEngine me = new MergeEngine())
                    {
                        string backupPath;
                        me.ApplyAndSave(path, items, backup, out backupPath);
                        backupOut = backupPath;
                    }
                    return true;
                },
                delegate(bool okk, Exception err) { error = err; });
            t.Join();
            if (error != null) throw error;
            return backupOut;
        }

        // ================================================================ 최근 비교(U1)
        private void BuildRecentMenu()
        {
            if (_btnRecent == null) return;
            _btnRecent.DropDownItems.Clear();
            List<string[]> pairs = _settings.GetRecentPairs();
            if (pairs.Count == 0)
            {
                ToolStripMenuItem empty = new ToolStripMenuItem("(최근 비교 없음)");
                empty.Enabled = false;
                _btnRecent.DropDownItems.Add(empty);
                return;
            }
            foreach (string[] pr in pairs)
            {
                string left = pr[0], right = pr[1];
                string text = FileLabel(left) + "  ↔  " + FileLabel(right);
                ToolStripMenuItem it = new ToolStripMenuItem(text);
                it.ToolTipText = "좌: " + left + "\r\n우: " + right;
                string cl = left, cr = right;   // 클로저 캡처(C#5 안전)
                it.Click += delegate { OpenRecentPair(cl, cr); };
                _btnRecent.DropDownItems.Add(it);
            }
        }

        private static string FileLabel(string path)
        {
            if (string.IsNullOrEmpty(path)) return "(없음)";
            try { return Path.GetFileName(path.TrimEnd('\\', '/')); }
            catch { return path; }
        }

        private void OpenRecentPair(string left, string right)
        {
            bool lok = File.Exists(left) || Directory.Exists(left);
            bool rok = File.Exists(right) || Directory.Exists(right);
            if (!lok || !rok)
            {
                Info("최근 비교 파일을 찾을 수 없습니다:\r\n"
                    + (lok ? "" : "· 좌: " + left + "\r\n")
                    + (rok ? "" : "· 우: " + right + "\r\n")
                    + "\r\n이동/삭제되었을 수 있습니다.");
                return;
            }
            _leftPath = left; _rightPath = right;
            _lblLeftPath.Text = left; _lblRightPath.Text = right;
            _leftWb = null; _rightWb = null;   // 캐시 무효화 후 재로드
            Logger.Info("최근 비교 선택 → 로드/비교: 좌=" + left + " 우=" + right);
            LoadAndCompare(IsCsvSession());
        }

        // ================================================================ 그리드 글자 크기(U5)
        private void ApplyGridFont(float size)
        {
            if (size < 6f) size = 6f;
            if (size > 22f) size = 22f;
            _gridFontSize = size;
            Font f = new Font(FontFamily.GenericSansSerif, size);
            int h = (int)Math.Ceiling(size * 1.7f) + 8;
            Font old = _gridFont;                     // 반복 조절 시 GDI HFONT 누적 방지
            ApplyGridFontOne(_gridLeft, f, h);
            ApplyGridFontOne(_gridRight, f, h);
            _gridFont = f;
            if (old != null) old.Dispose();
            _settings.GridFontSize = size;   // 저장은 종료 시(SaveWindow)
        }

        private void ApplyGridFontOne(DataGridView g, Font f, int h)
        {
            if (g == null) return;
            g.DefaultCellStyle.Font = f;
            g.RowTemplate.Height = h;
            // 가상모드: 이미 만들어진 행에도 새 높이를 반영하려면 RowCount 를 재설정.
            if (g.Columns.Count > 0)
            {
                int keep = g.RowCount;
                g.RowCount = 0;
                g.RowCount = keep;
            }
        }

        private void Grid_MouseWheel(object sender, MouseEventArgs e)
        {
            if ((ModifierKeys & Keys.Control) != Keys.Control) return;
            ApplyGridFont(_gridFontSize + (e.Delta > 0 ? 1f : -1f));
            HandledMouseEventArgs he = e as HandledMouseEventArgs;
            if (he != null) he.Handled = true;   // 스크롤 대신 확대/축소로 소비
        }

        // ================================================================ 열 자동맞춤(U9)
        private void AutoFitColumns()
        {
            try
            {
                _gridLeft.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
                _gridRight.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.DisplayedCells);
            }
            catch { }
        }

        private void Grid_ColHeaderDoubleClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.ColumnIndex < 0) return;
            try
            {
                // 좌우 정렬 유지를 위해 같은 열을 양쪽에 동일 적용.
                if (e.ColumnIndex < _gridLeft.ColumnCount)
                    _gridLeft.AutoResizeColumn(e.ColumnIndex, DataGridViewAutoSizeColumnMode.DisplayedCells);
                if (e.ColumnIndex < _gridRight.ColumnCount)
                    _gridRight.AutoResizeColumn(e.ColumnIndex, DataGridViewAutoSizeColumnMode.DisplayedCells);
            }
            catch { }
        }

        // ================================================================ 병합 취소(U8)
        private void CancelAllMerges()
        {
            int n = _pendingLeft.Count + _pendingRight.Count;
            if (n == 0) { Info("취소할 병합 대기가 없습니다."); return; }
            DialogResult ans = MessageBox.Show(this,
                "병합 대기 중인 변경 " + n + "건을 모두 취소합니다.\r\n계속하시겠어요?",
                "병합 취소", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (ans != DialogResult.Yes) return;
            _pendingLeft.Clear();
            _pendingRight.Clear();
            _gridLeft.Invalidate();
            _gridRight.Invalidate();
            _lblSummary.Text = "병합 대기를 모두 취소했습니다.";
            Logger.Info("병합 대기 전체 취소 (" + n + "건)");
            UpdateButtonStates();
        }

        private void Grid_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (e.Button != MouseButtons.Right) return;
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
            DataGridView g = (DataGridView)sender;
            _ctxGrid = g; _ctxRow = e.RowIndex; _ctxCol = e.ColumnIndex;
            // 우클릭 셀을 현재 셀로(선택 미러링 포함).
            try { g.CurrentCell = g.Rows[e.RowIndex].Cells[e.ColumnIndex]; }
            catch { }
            _cellMenuCancel.Enabled = IsCellPending(g, e.RowIndex, e.ColumnIndex);
            _cellMenu.Show(g, g.PointToClient(Cursor.Position));
        }

        private bool IsCellPending(DataGridView g, int row, int col)
        {
            if (_curSheet == null || row < 0 || row >= _rowMap.Length || col < 0) return false;
            DiffRow dr = _curSheet.Rows[_rowMap[row]];
            bool isLeft = (g == _gridLeft);
            int absRow = isLeft ? dr.LeftRow : dr.RightRow;
            if (absRow < 0) return false;
            int absCol = _colStart + col;
            string key = CellKey(_curSheet.Name, absRow, absCol);
            return (isLeft ? _pendingLeft : _pendingRight).ContainsKey(key);
        }

        private void CancelOneCell()
        {
            if (_ctxGrid == null || _curSheet == null) return;
            if (_ctxRow < 0 || _ctxRow >= _rowMap.Length || _ctxCol < 0) return;
            DiffRow dr = _curSheet.Rows[_rowMap[_ctxRow]];
            bool isLeft = (_ctxGrid == _gridLeft);
            int absRow = isLeft ? dr.LeftRow : dr.RightRow;
            if (absRow < 0) return;
            int absCol = _colStart + _ctxCol;
            string key = CellKey(_curSheet.Name, absRow, absCol);
            Dictionary<string, object> pend = isLeft ? _pendingLeft : _pendingRight;
            if (pend.Remove(key))
            {
                _gridLeft.Invalidate();
                _gridRight.Invalidate();
                _lblSummary.Text = "셀 병합 취소: " + ColLetter(absCol) + absRow
                    + "  (대기 좌:" + _pendingLeft.Count + " 우:" + _pendingRight.Count + ")";
                Logger.Info("셀 병합 취소 " + (isLeft ? "좌" : "우") + " " + key);
                UpdateButtonStates();
            }
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
                Logger.Info("드래그앤드롭 2파일 → 로드/비교");
                LoadAndCompare(false);
            }
            else
            {
                bool left = string.IsNullOrEmpty(_leftPath);
                if (left) { _leftPath = files[0]; _lblLeftPath.Text = files[0]; }
                else { _rightPath = files[0]; _lblRightPath.Text = files[0]; }
                Logger.Info("드래그앤드롭 1파일(" + (left ? "좌" : "우") + ") → 미리보기");
                LoadOneAsync(files[0], left);
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
