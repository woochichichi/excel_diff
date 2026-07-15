using System;

namespace ExcelDiffMerge
{
    /// <summary>
    /// 워크북 로더 추상화. 운영은 ExcelComReader(Excel COM), 로컬 테스트는 CsvWorkbookReader.
    /// 덕분에 Excel/DRM 없이도 diff 엔진/UI 로직을 로컬에서 검증할 수 있다.
    /// </summary>
    public interface IWorkbookReader : IDisposable
    {
        WorkbookData LoadWorkbook(string path);
    }
}
