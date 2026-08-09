using System;
using Excel = Microsoft.Office.Interop.Excel;

namespace Fuwoa.AddIn
{
    /// <summary>
    /// Excel 操作保护上下文：进入时关闭屏幕刷新/自动计算/事件，退出时恢复。
    /// </summary>
    internal sealed class ExcelGuard : IDisposable
    {
        private readonly Excel.Application _app;
        private readonly bool _prevScreenUpdating;
        private readonly Excel.XlCalculation _prevCalc;
        private readonly bool _prevEvents;
        private readonly string _prevStatusBar;

        public ExcelGuard(Excel.Application app)
        {
            _app = app ?? throw new ArgumentNullException(nameof(app));
            try { _prevScreenUpdating = app.ScreenUpdating; } catch { _prevScreenUpdating = true; }
            try { _prevCalc = app.Calculation; } catch { _prevCalc = Excel.XlCalculation.xlCalculationAutomatic; }
            try { _prevEvents = app.EnableEvents; } catch { _prevEvents = true; }
            try { _prevStatusBar = app.StatusBar as string ?? ""; } catch { _prevStatusBar = ""; }

            try { app.ScreenUpdating = false; } catch { }
            try { app.Calculation = Excel.XlCalculation.xlCalculationManual; } catch { }
            try { app.EnableEvents = false; } catch { }
            try { app.StatusBar = ""; } catch { }
        }

        public void Dispose()
        {
            try { _app.ScreenUpdating = _prevScreenUpdating; } catch { }
            try { _app.Calculation = _prevCalc; } catch { }
            try { _app.EnableEvents = _prevEvents; } catch { }
            try { _app.StatusBar = _prevStatusBar; } catch { }
            try { _app.Calculate(); } catch { }
        }
    }
}
