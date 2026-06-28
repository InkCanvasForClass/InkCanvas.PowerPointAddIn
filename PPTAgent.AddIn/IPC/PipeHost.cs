using System;
using System.IO;
using System.IO.Pipes;
using System.Threading;
using System.Threading.Tasks;
using PPTAgent.Contracts;

namespace PPTAgent.AddIn.IPC
{
    public sealed class PipeHost : IDisposable
    {
        private readonly Func<string, string> _dispatch;
        private readonly object _sendLock = new object();
        private CancellationTokenSource _cts;
        private NamedPipeServerStream _currentPipe;
        private volatile bool _clientConnected;
        private bool _disposed;

        public bool IsEnabled => _clientConnected;

        public PipeHost(Func<string, string> dispatch)
        {
            _dispatch = dispatch;
        }

        /// <summary>
        /// Push a message (state/event) to the currently connected client.
        /// Thread-safe, can be called from any thread.
        /// </summary>
        public void SendFrame(string json)
        {
            var pipe = _currentPipe;
            if (pipe == null || !pipe.IsConnected || string.IsNullOrEmpty(json)) return;

            lock (_sendLock)
            {
                try
                {
                    if (pipe.IsConnected)
                        PipeFrame.WriteFrame(pipe, json);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"PPTAgent push error: {ex.Message}");
                }
            }
        }

        public void Start()
        {
            if (_cts != null) return;
            _cts = new CancellationTokenSource();
            Task.Run(() => AcceptLoop(_cts.Token));
        }

        public void Stop()
        {
            try { _cts?.Cancel(); } catch { }
            _cts?.Dispose();
            _cts = null;
            _clientConnected = false;
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            Stop();
        }

        private async Task AcceptLoop(CancellationToken token)
        {
            while (!token.IsCancellationRequested)
            {
                NamedPipeServerStream pipe = null;
                try
                {
                    pipe = new NamedPipeServerStream(
                        PipeConstants.PipeName,
                        PipeDirection.InOut,
                        1,
                        PipeTransmissionMode.Byte,
                        PipeOptions.Asynchronous);

                    await pipe.WaitForConnectionAsync(token).ConfigureAwait(false);
                    _currentPipe = pipe;
                    _clientConnected = true;
                    System.Diagnostics.Debug.WriteLine("PPTAgent: client connected");

                    HandleClient(pipe, token);
                }
                catch (OperationCanceledException) { break; }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"PPTAgent pipe error: {ex.Message}");
                }
                finally
                {
                    _clientConnected = false;
                    _currentPipe = null;
                    try { pipe?.Dispose(); } catch { }
                }
            }
        }

        private void HandleClient(NamedPipeServerStream pipe, CancellationToken token)
        {
            while (!token.IsCancellationRequested && pipe.IsConnected)
            {
                try
                {
                    string requestJson = PipeFrame.ReadFrame(pipe);
                    string responseJson = _dispatch.Invoke(requestJson);
                    if (!string.IsNullOrEmpty(responseJson))
                    {
                        lock (_sendLock)
                        {
                            if (pipe.IsConnected)
                                PipeFrame.WriteFrame(pipe, responseJson);
                        }
                    }
                }
                catch (IOException) { break; }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"PPTAgent handle error: {ex.Message}");
                    break;
                }
            }
        }
    }
}
