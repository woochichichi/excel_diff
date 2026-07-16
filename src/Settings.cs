using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 아주 단순한 key=value ini 저장/복원 (spec §6 편의 요소).
    /// 외부 라이브러리 0 원칙 → 자체 파서. exe 옆 settings.ini 사용.
    /// 최근 폴더/최근 비교 이력/온보딩 표시 여부 등을 기억.
    /// </summary>
    public sealed class Settings
    {
        private readonly string _path;
        private readonly Dictionary<string, string> _map =
            new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public Settings(string path)
        {
            _path = path;
            Load();
        }

        public static Settings LoadDefault()
        {
            string exeDir = AppDomain.CurrentDomain.BaseDirectory;
            return new Settings(Path.Combine(exeDir, "settings.ini"));
        }

        public string Get(string key, string def)
        {
            string v;
            return _map.TryGetValue(key, out v) ? v : def;
        }

        public bool GetBool(string key, bool def)
        {
            string v = Get(key, null);
            if (v == null) return def;
            bool b;
            return bool.TryParse(v, out b) ? b : def;
        }

        public void Set(string key, string value)
        {
            _map[key] = value ?? "";
        }

        public void SetBool(string key, bool value)
        {
            _map[key] = value ? "true" : "false";
        }

        // 최근 폴더.
        public string LastFolder
        {
            get { return Get("LastFolder", ""); }
            set { Set("LastFolder", value); }
        }

        // 온보딩 다시 보지 않기.
        public bool HideOnboarding
        {
            get { return GetBool("HideOnboarding", false); }
            set { SetBool("HideOnboarding", value); }
        }

        // 그리드 글자 크기(양쪽 그리드 공통). 0 이하면 '미설정'.
        public float GridFontSize
        {
            get
            {
                float v;
                if (float.TryParse(Get("GridFontSize", ""), NumberStyles.Float, CultureInfo.InvariantCulture, out v))
                    return v;
                return 0f;
            }
            set { Set("GridFontSize", value.ToString(CultureInfo.InvariantCulture)); }
        }

        // 최근 비교 이력(최대 10). "left|right" 형태를 개행으로 구분해 저장.
        public List<string[]> GetRecentPairs()
        {
            List<string[]> list = new List<string[]>();
            for (int i = 0; i < 10; i++)
            {
                string v = Get("Recent" + i, null);
                if (string.IsNullOrEmpty(v)) continue;
                int bar = v.IndexOf('|');
                if (bar <= 0) continue;
                list.Add(new string[] { v.Substring(0, bar), v.Substring(bar + 1) });
            }
            return list;
        }

        public void AddRecentPair(string left, string right)
        {
            List<string[]> list = GetRecentPairs();
            // 중복 제거(맨 앞으로).
            list.RemoveAll(delegate(string[] p)
            {
                return string.Equals(p[0], left, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(p[1], right, StringComparison.OrdinalIgnoreCase);
            });
            list.Insert(0, new string[] { left, right });
            if (list.Count > 10) list.RemoveRange(10, list.Count - 10);
            for (int i = 0; i < 10; i++)
            {
                if (i < list.Count) Set("Recent" + i, list[i][0] + "|" + list[i][1]);
                else Set("Recent" + i, "");
            }
        }

        private void Load()
        {
            _map.Clear();
            if (!File.Exists(_path)) return;
            try
            {
                foreach (string raw in File.ReadAllLines(_path, Encoding.UTF8))
                {
                    string line = raw.Trim();
                    if (line.Length == 0 || line[0] == '#' || line[0] == ';' || line[0] == '[') continue;
                    int eq = line.IndexOf('=');
                    if (eq <= 0) continue;
                    string k = line.Substring(0, eq).Trim();
                    string v = line.Substring(eq + 1).Trim();
                    _map[k] = v;
                }
            }
            catch { /* 설정 로드 실패는 치명적이지 않음 */ }
        }

        public void Save()
        {
            try
            {
                StringBuilder sb = new StringBuilder();
                sb.AppendLine("# ExcelDiffMerge settings");
                foreach (KeyValuePair<string, string> kv in _map)
                    sb.Append(kv.Key).Append('=').Append(kv.Value).Append("\r\n");
                File.WriteAllText(_path, sb.ToString(), Encoding.UTF8);
            }
            catch { /* 저장 실패 무시 */ }
        }
    }
}
