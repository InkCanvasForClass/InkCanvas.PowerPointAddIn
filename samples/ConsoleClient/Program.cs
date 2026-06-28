using System;
using System.IO.Pipes;
using System.Threading;
using Newtonsoft.Json;
using PPTAgent.Contracts;

namespace PPTAgent.Samples.ConsoleClient
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("PPTAgent Console Client");
            Console.WriteLine("=======================");
            Console.WriteLine($"Connecting to pipe: {PipeConstants.PipeName}");
            Console.WriteLine();

            try
            {
                using var pipe = new NamedPipeClientStream(".", PipeConstants.PipeName, PipeDirection.InOut);
                Console.Write("Connecting...");
                pipe.Connect(5000);
                Console.WriteLine(" Connected!");
                Console.WriteLine();

                // Start background listener for events
                var cts = new CancellationTokenSource();
                var listener = new Thread(() => ListenForEvents(pipe, cts.Token))
                {
                    IsBackground = true
                };
                listener.Start();

                // Interactive command loop
                while (true)
                {
                    Console.WriteLine("Commands:");
                    Console.WriteLine("  1) Get state");
                    Console.WriteLine("  2) Next slide");
                    Console.WriteLine("  3) Previous slide");
                    Console.WriteLine("  4) Goto slide");
                    Console.WriteLine("  5) Start slide show");
                    Console.WriteLine("  6) End slide show");
                    Console.WriteLine("  7) Show slide navigation");
                    Console.WriteLine("  8) Disable auto-play timings");
                    Console.WriteLine("  9) Unhide hidden slides");
                    Console.WriteLine("  0) Exit");
                    Console.Write("> ");

                    var key = Console.ReadLine()?.Trim();
                    Console.WriteLine();

                    switch (key)
                    {
                        case "1":
                            SendCommand(pipe, PPTCommands.State);
                            break;
                        case "2":
                            SendCommand(pipe, PPTCommands.Next);
                            break;
                        case "3":
                            SendCommand(pipe, PPTCommands.Previous);
                            break;
                        case "4":
                            Console.Write("Slide number: ");
                            if (int.TryParse(Console.ReadLine(), out int num))
                            {
                                SendCommand(pipe, PPTCommands.GotoSlide, new GotoSlideRequest { SlideNumber = num });
                            }
                            break;
                        case "5":
                            SendCommand(pipe, PPTCommands.StartSlideShow);
                            break;
                        case "6":
                            SendCommand(pipe, PPTCommands.EndSlideShow);
                            break;
                        case "7":
                            SendCommand(pipe, PPTCommands.ShowSlideNavigation);
                            break;
                        case "8":
                            SendCommand(pipe, PPTCommands.DisableAutoPlayTimings);
                            break;
                        case "9":
                            SendCommand(pipe, PPTCommands.UnhideHiddenSlides);
                            break;
                        case "0":
                            cts.Cancel();
                            return;
                        default:
                            Console.WriteLine("Invalid option.");
                            break;
                    }

                    Console.WriteLine();
                }
            }
            catch (TimeoutException)
            {
                Console.WriteLine("Failed to connect. Is PowerPoint running with PPTAgent add-in loaded?");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error: {ex.Message}");
            }
        }

        static void SendCommand(NamedPipeClientStream pipe, string command, object data = null)
        {
            var requestId = Guid.NewGuid().ToString("N");
            var message = new PPTPipeMessage<object>
            {
                Version = PipeConstants.ProtocolVersion,
                Type = PPTMessageTypes.Command,
                Cmd = command,
                Data = data,
                RequestId = requestId
            };

            var json = JsonConvert.SerializeObject(message);
            PipeFrame.WriteFrame(pipe, json);

            // Read response
            var responseJson = PipeFrame.ReadFrame(pipe);
            var response = JsonConvert.DeserializeObject<PPTPipeMessage<object>>(responseJson);

            if (response.Success)
            {
                Console.WriteLine($"[{command}] Success");
                if (response.Data != null)
                {
                    Console.WriteLine(JsonConvert.SerializeObject(response.Data, Formatting.Indented));
                }
            }
            else
            {
                Console.WriteLine($"[{command}] Error: {response.Error}");
            }
        }

        static void ListenForEvents(NamedPipeClientStream pipe, CancellationToken token)
        {
            try
            {
                while (!token.IsCancellationRequested && pipe.IsConnected)
                {
                    var json = PipeFrame.ReadFrame(pipe);
                    var msg = JsonConvert.DeserializeObject<PPTPipeMessage<object>>(json);

                    if (msg.Type == PPTMessageTypes.Event)
                    {
                        Console.WriteLine();
                        Console.WriteLine($"[Event] {msg.Cmd}");
                        if (msg.Data != null)
                        {
                            Console.WriteLine(JsonConvert.SerializeObject(msg.Data, Formatting.Indented));
                        }
                        Console.Write("> ");
                    }
                }
            }
            catch
            {
                // Pipe closed or error
            }
        }
    }
}
