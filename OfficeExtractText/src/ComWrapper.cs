using System;
using System.Runtime.InteropServices;

namespace OfficeExtractTexts
{
    class ComWrapper<T> : IDisposable
    {
        public T ComObject { get; }

        public ComWrapper(T comObject)
        {
            this.ComObject = comObject;
        }

        private bool disposedValue = false;

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue)
            {
                if (disposing)
                {
                    //nop
                }
                Marshal.ReleaseComObject(ComObject);
                disposedValue = true;
            }
        }

        ~ComWrapper()
        {
            Dispose(false);
        }

        public void Dispose()
        {
            Dispose(true);
        }

    }
}
