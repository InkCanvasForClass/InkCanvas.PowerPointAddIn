# PPTAgent

A VSTO add-in that turns Microsoft PowerPoint into a remote-controlled service via Windows Named Pipes.

Any application can connect to the pipe and control PowerPoint — slide navigation, slide show management, media region detection, and more.

## Architecture

```
┌─────────────────────┐     Named Pipe      ┌──────────────────┐
│  Any Application    │ ◄──── JSON ────►    │  PowerPoint      │
│  (Console, WPF,     │   "PPTAgent_Pipe"   │  + PPTAgent VSTO │
│   Python, C++, ...) │                     │  Add-in          │
└─────────────────────┘                     └──────────────────┘
```

## Projects

| Project | Target | Description |
|---------|--------|-------------|
| `PPTAgent.Contracts` | netstandard2.0 | Protocol definitions — message types, pipe frame format, state models. **This is the SDK.** |
| `PPTAgent.AddIn` | net472 (VSTO) | The VSTO add-in that runs inside PowerPoint. Requires Visual Studio with Office development tools. |
| `samples/ConsoleClient` | net8.0 | Sample interactive client demonstrating the protocol. |

## Quick Start

### 1. Build

```bash
dotnet build PPTAgent.sln
```

### 2. Install the Add-in

Deploy `PPTAgent.AddIn.vsto` (found in the build output) via:
- Double-click the `.vsto` file, or
- Use `InstallUtil.exe`, or
- Register via Visual Studio

### 3. Connect from Your App

**Option A: Reference `PPTAgent.Contracts` NuGet / DLL**

```csharp
using System.IO.Pipes;
using PPTAgent.Contracts;
using Newtonsoft.Json;

var pipe = new NamedPipeClientStream(".", PipeConstants.PipeName, PipeDirection.InOut);
pipe.Connect();

// Send a command
var msg = new PPTPipeMessage<object>
{
    Version = 1,
    Type = "cmd",
    Cmd = "next",
    RequestId = Guid.NewGuid().ToString("N")
};
PipeFrame.WriteFrame(pipe, JsonConvert.SerializeObject(msg));

// Read response
string json = PipeFrame.ReadFrame(pipe);
var response = JsonConvert.DeserializeObject<PPTPipeMessage<object>>(json);
```

**Option B: No dependencies (any language)**

The wire protocol is trivial:

```
[4-byte LE int: payload length][UTF-8 JSON payload]
```

JSON message format:

```json
{
    "Version": 1,
    "Type": "cmd",
    "Cmd": "next",
    "Data": null,
    "RequestId": "unique-id"
}
```

## Protocol Reference

### Commands (Client → Add-in)

| Cmd | Data | Description |
|-----|------|-------------|
| `state` | — | Get current PowerPoint state |
| `next` | — | Next slide |
| `prev` | — | Previous slide |
| `gotoSlide` | `{"SlideNumber": N}` | Jump to slide N |
| `startSlideShow` | — | Start slide show |
| `endSlideShow` | — | End slide show |
| `showSlideNavigation` | — | Show slide navigation panel |
| `disableAutoPlayTimings` | — | Disable auto-advance |
| `unhideHiddenSlides` | — | Unhide all hidden slides |
| `getMediaRegions` | — | Get media control regions on current slide |
| `exportSlideThumbnails` | `{"Width":W,"Height":H}` | Export slide thumbnails (not yet implemented) |

### Response (Add-in → Client)

```json
{
    "Version": 1,
    "Type": "response",
    "Cmd": "state",
    "Success": true,
    "Data": { "SlideIndex": 1, "TotalSlides": 10, "IsRunning": true, ... },
    "RequestId": "matching-request-id"
}
```

### Event Push (Add-in → Client, unsolicited)

```json
{
    "Version": 1,
    "Type": "event",
    "Cmd": "slideShowBegin",
    "Data": { "SlideIndex": 1, "TotalSlides": 10, "IsRunning": true, ... }
}
```

Events: `presentationOpen`, `presentationClose`, `slideShowBegin`, `slideShowNextSlide`, `slideShowEnd`

### Error Response

```json
{
    "Version": 1,
    "Type": "error",
    "Success": false,
    "Error": "Unknown command: foo",
    "RequestId": "matching-request-id"
}
```

## Configuration

Change the pipe name before starting the host (e.g., for multi-instance scenarios):

```csharp
PipeConstants.PipeName = "MyCustomPipeName";
```

## Sample Client

Run the included console client:

```bash
cd samples/ConsoleClient
dotnet run
```

It provides an interactive menu to send commands and displays events in real-time.

## License

GPL-3.0
