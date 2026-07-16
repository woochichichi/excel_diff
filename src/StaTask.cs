using System;
using System.Threading;

namespace ExcelDiffMerge
{
    /// <summary>
    /// Excel COM 작업은 반드시 STA 스레드에서 실행해야 한다(정찰/조사 반영).
    /// BackgroundWorker/ThreadPool 은 MTA 라 간헐적 RPC/마샬링 오류가 난다.
    /// 이 헬퍼는 전용 STA 스레드에서 작업을 돌리고 결과/예외를 UI 스레드로 전달한다.
    /// (UI 블로킹 방지는 호출측에서 별도 STA 스레드로 띄우고 완료 콜백을 Control.Invoke 로 처리)
    /// </summary>
    public static class StaTask
    {
        /// <summary>func 를 전용 STA 스레드에서 실행하고, 완료 시 onDone(result, error) 를 호출한다.
        /// onDone 은 STA 워커 스레드에서 호출되므로, UI 갱신은 콜백 안에서 Control.Invoke 로 마샬링할 것.</summary>
        public static Thread Run<T>(Func<T> func, Action<T, Exception> onDone)
        {
            Thread t = new Thread(delegate()
            {
                T result = default(T);
                Exception error = null;
                try { result = func(); }
                catch (Exception ex) { error = ex; }
                // onDone 은 보통 폼의 BeginInvoke 로 UI 스레드에 마샬링한다.
                // 로드 중 사용자가 폼을 닫으면 폼/핸들이 사라져 ObjectDisposedException/
                // InvalidOperationException 이 STA 워커의 '미처리 예외'가 되어 프로세스가 죽는다.
                // → 폼이 이미 사라진 정상 종료 경로이므로 로깅만 하고 무시한다(B1).
                if (onDone != null)
                {
                    try { onDone(result, error); }
                    catch (Exception exDone)
                    {
                        try { Logger.Error("완료 콜백 실행 실패(폼이 닫혔을 수 있음, 무시)", exDone); }
                        catch { }
                    }
                }
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            return t;
        }
    }
}
