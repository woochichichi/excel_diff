using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 스크롤바 옆 변경 밀도 마커(조사 UX 반영). 차이 행 위치에 눈금을 그려
    /// 큰 시트에서 변경이 어디 몰려있는지 한눈에 보여주고, 클릭하면 그 위치로 점프한다.
    /// </summary>
    public sealed class MarkerBar : Control
    {
        // 0..1 정규화 위치 + 색상.
        private readonly List<KeyValuePair<float, Color>> _marks = new List<KeyValuePair<float, Color>>();
        private float _viewport = -1f; // 현재 보이는 위치(0..1), 음수면 숨김

        /// <summary>클릭 시 호출(0..1 위치).</summary>
        public Action<float> OnSeek;

        public MarkerBar()
        {
            Width = 14;
            DoubleBuffered = true;
            SetStyle(ControlStyles.ResizeRedraw, true);
        }

        public void SetMarks(List<KeyValuePair<float, Color>> marks)
        {
            _marks.Clear();
            if (marks != null) _marks.AddRange(marks);
            Invalidate();
        }

        public void SetViewport(float pos)
        {
            _viewport = pos;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            Graphics g = e.Graphics;
            g.Clear(Color.FromArgb(245, 245, 245));
            int h = Height;
            foreach (KeyValuePair<float, Color> m in _marks)
            {
                int y = (int)(m.Key * (h - 1));
                using (Pen p = new Pen(m.Value, 2))
                    g.DrawLine(p, 2, y, Width - 2, y);
            }
            if (_viewport >= 0f)
            {
                int y = (int)(_viewport * (h - 1));
                using (Pen p = new Pen(Color.FromArgb(90, 90, 90), 1))
                    g.DrawLine(p, 0, y, Width, y);
            }
            using (Pen border = new Pen(Color.FromArgb(200, 200, 200)))
                g.DrawRectangle(border, 0, 0, Width - 1, Height - 1);
        }

        protected override void OnMouseDown(MouseEventArgs e)
        {
            base.OnMouseDown(e);
            if (OnSeek != null && Height > 1)
            {
                float pos = (float)e.Y / (Height - 1);
                if (pos < 0f) pos = 0f;
                if (pos > 1f) pos = 1f;
                OnSeek(pos);
            }
        }
    }
}
