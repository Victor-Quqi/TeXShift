using System;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using TeXShift.Core.Services;
using TeXShift.Core.Logging;

namespace TeXShift.Core.Tests
{
    [TestClass]
    public class StaWorkerTests
    {
        private static readonly Type WorkerType = typeof(ServiceContainer).Assembly.GetType("TeXShift.Core.Utils.StaWorker", true);

        private static IDisposable Create(Action cleanup) =>
            (IDisposable)Activator.CreateInstance(WorkerType, "TeXShift_Test_STA", cleanup);

        private static Task<T> Run<T>(IDisposable worker, Func<Task<T>> action) =>
            (Task<T>)WorkerType.GetMethod("InvokeAsync").MakeGenericMethod(typeof(T)).Invoke(worker, new object[] { action });

        [TestMethod]
        public async Task PerformanceTraceFlowsToStaAndRestoresNestedSession()
        {
            var current = typeof(PerformanceTraceContext).GetMethod("GetCurrent", BindingFlags.Static | BindingFlags.NonPublic);
            Assert.IsNull(current.Invoke(null, null));
            using (var worker = Create(() => { }))
            using (PerformanceTraceContext.BeginCommand("test", "outer"))
            {
                var outer = current.Invoke(null, null);
                await Run(worker, async () =>
                {
                    Assert.AreSame(outer, current.Invoke(null, null));
                    using (PerformanceTraceContext.Measure("STA work"))
                    {
                        await Task.Yield();
                        Assert.AreSame(outer, current.Invoke(null, null));
                    }
                    return true;
                });
                using (PerformanceTraceContext.BeginCommand("test", "inner"))
                    Assert.AreNotSame(outer, current.Invoke(null, null));
                Assert.AreSame(outer, current.Invoke(null, null));
                string report = (string)outer.GetType().GetMethod("FormatReport").Invoke(outer, null);
                StringAssert.Contains(report, "STA work");
            }
            Assert.IsNull(current.Invoke(null, null));
        }

        [TestMethod]
        public async Task DisposeWaitsForCleanupOnOwningSta()
        {
            Thread owner = null;
            int cleaned = 0;
            using (var worker = Create(() =>
            {
                Assert.AreSame(owner, Thread.CurrentThread);
                Interlocked.Increment(ref cleaned);
            }))
            {
                owner = await Run(worker, async () =>
                {
                    var before = Thread.CurrentThread;
                    await Task.Delay(10);
                    Assert.AreSame(before, Thread.CurrentThread);
                    Assert.AreEqual(ApartmentState.STA, before.GetApartmentState());
                    return before;
                });
                worker.Dispose();
                worker.Dispose();
                Assert.AreEqual(1, cleaned);
                Assert.IsFalse(owner.IsAlive);
            }
        }

        [TestMethod]
        public async Task DisposeCompletesPendingOperationAndRejectsNewWork()
        {
            using (var worker = Create(() => { }))
            using (var started = new ManualResetEventSlim())
            {
                var pending = Run(worker, () =>
                {
                    started.Set();
                    return new TaskCompletionSource<bool>().Task;
                });
                Assert.IsTrue(started.Wait(TimeSpan.FromSeconds(5)));
                worker.Dispose();
                Assert.AreSame(pending, await Task.WhenAny(pending, Task.Delay(5000)));
                Assert.IsTrue(pending.IsCanceled);
                try
                {
                    _ = Run(worker, () => Task.FromResult(true));
                    Assert.Fail("Work was accepted after disposal.");
                }
                catch (TargetInvocationException ex)
                {
                    Assert.IsInstanceOfType<ObjectDisposedException>(ex.InnerException);
                }
            }
        }

        [TestMethod]
        public void CleanupFailureStillStopsThread()
        {
            Thread owner = null;
            var worker = Create(() => { throw new InvalidOperationException("cleanup failed"); });
            owner = Run(worker, () => Task.FromResult(Thread.CurrentThread)).GetAwaiter().GetResult();
            try
            {
                worker.Dispose();
                Assert.Fail("Cleanup failure was lost.");
            }
            catch (InvalidOperationException ex)
            {
                Assert.AreEqual("cleanup failed", ex.InnerException.Message);
                Assert.IsFalse(owner.IsAlive);
            }
        }
    }
}
