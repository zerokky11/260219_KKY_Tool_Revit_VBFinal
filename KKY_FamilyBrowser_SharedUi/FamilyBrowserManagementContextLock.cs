using System;
using System.IO;
using System.Threading;

public static class FamilyBrowserManagementContextLock
{
    private const string MutexName = "Local\\KKY_FamilyBrowser_ManagementContext_v1";

    public static IDisposable Acquire(TimeSpan timeout)
    {
        Mutex mutex = new Mutex(false, MutexName);
        bool acquired = false;
        try
        {
            try
            {
                acquired = mutex.WaitOne(timeout);
            }
            catch (AbandonedMutexException)
            {
                acquired = true;
            }
            if (!acquired)
            {
                throw new IOException("Timed out waiting for another Revit process to finish changing the Family Browser management context.");
            }
            return new Lease(mutex);
        }
        catch
        {
            mutex.Dispose();
            throw;
        }
    }

    private sealed class Lease : IDisposable
    {
        private Mutex _mutex;

        public Lease(Mutex mutex)
        {
            _mutex = mutex;
        }

        public void Dispose()
        {
            Mutex mutex = Interlocked.Exchange(ref _mutex, null);
            if (mutex == null)
            {
                return;
            }
            try
            {
                mutex.ReleaseMutex();
            }
            finally
            {
                mutex.Dispose();
            }
        }
    }
}
