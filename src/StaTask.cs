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
                if (onDone != null) onDone(result, error);
            });
            t.IsBackground = true;
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            return t;
        }
    }
}
