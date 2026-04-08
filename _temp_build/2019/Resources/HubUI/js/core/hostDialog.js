export function refreshUiAfterHostDialog(render, delay = 120) {
  if (typeof render !== 'function') return;

  const run = () => {
    try { render(); } catch { }
  };

  run();

  if (typeof window === 'undefined') return;

  let released = false;
  let raf1 = 0;
  let raf2 = 0;
  let timerRun = 0;
  let timerFinalize = 0;

  const cleanup = () => {
    if (released) return;
    released = true;
    window.removeEventListener('focus', onFocus, true);
    if (typeof document !== 'undefined') {
      document.removeEventListener('visibilitychange', onVisible, true);
    }
    if (raf1 && typeof window.cancelAnimationFrame === 'function') window.cancelAnimationFrame(raf1);
    if (raf2 && typeof window.cancelAnimationFrame === 'function') window.cancelAnimationFrame(raf2);
    if (timerRun) window.clearTimeout(timerRun);
    if (timerFinalize) window.clearTimeout(timerFinalize);
  };

  const rerender = () => {
    cleanup();
    run();
  };

  const onFocus = () => rerender();
  const onVisible = () => {
    if (document.visibilityState === 'visible') rerender();
  };

  window.addEventListener('focus', onFocus, true);
  if (typeof document !== 'undefined') {
    document.addEventListener('visibilitychange', onVisible, true);
  }

  if (typeof window.requestAnimationFrame === 'function') {
    raf1 = window.requestAnimationFrame(() => {
      run();
      raf2 = window.requestAnimationFrame(run);
    });
  } else {
    timerRun = window.setTimeout(run, 0);
  }

  timerFinalize = window.setTimeout(rerender, Math.max(0, Number(delay) || 0));
}
