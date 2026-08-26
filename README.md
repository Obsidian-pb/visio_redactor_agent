**English** | [Русский](README.ru.md)

# Visio Smart Redactor

**Visio Smart Redactor** is a Microsoft Visio stencil that integrates an LLM agent directly into the Visio environment. The agent can read and modify diagrams: draw shapes, edit text, format cells, work with files — all through natural language in a built-in chat.

The project is built on top of the [VBA-Harness](https://github.com/Obsidian-pb/vba_harness) library and uses [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) for JSON handling.

---

## Table of Contents

- [Features](#features)
- [Architecture](#architecture)
- [Agent Tools](#agent-tools)
- [Installation](#installation)
- [Configuration](#configuration)
- [Usage](#usage)
- [Project Structure](#project-structure)
- [Developing Custom Tools](#developing-custom-tools)
- [Dependencies](#dependencies)
- [License](#license)

---

## Features

- **LLM chat inside Visio** — a built-in chat window based on HTML rendering (WebBrowser control) with support for reasoning, tool-calling, and a work indicator.
- **Tool calling (function calling)** — the agent automatically calls tools to interact with the Visio document and the file system.
- **Multi-turn agent loop** — the agent can sequentially call multiple tools before giving a final answer (up to 30 turns by default).
- **Document context (RAG)** — text file contents can be added to the LLM context as reference material.
- **Model switching** — choose between a "smart" and a "fast" model directly in the chat.
- **Windows registry settings** — API URL, key, models, and system prompt are stored in `HKCU\Software\LLMVisio\`.
- **32/64-bit Office compatibility** — Win32 API declarations are adapted for both bitnesses.

---

## Architecture

The project follows the **Harness** pattern — a single orchestrator class that manages the LLM communication loop:

```
User → LLM_chat_html (chat form)
                    ↓
              clsHarness (agent loop)
                    ↓
         HTTP request to LLM API (OpenAI-compatible)
                    ↓
         Parse response (tool_calls / content / reasoning)
                    ↓
      clsHarnessToolsHandler → tool dispatcher
                    ↓
        clsTool_* (concrete Visio tools)
                    ↓
         Result returned to message history
                    ↓
            Loop repeats until LLM produces
            a final text response
```

### Key Components

| Component | Purpose |
|---|---|
| `clsHarness` | Main agent class: stores message history, sends HTTP requests to the LLM, parses responses, manages the tool-calling loop. |
| `clsHarnessToolsHandler` | Tool registry: registers all `clsTool_*`, generates the `tools[]` JSON schema for the LLM, dispatches calls by name. |
| `clsToolBase` | Abstract base class for all tools. Each tool implements `Name`, `Description`, `Schema` properties and the `Execute` method. |
| `clsTool_*` | Concrete tools for working with Visio and files (see [below](#agent-tools)). |
| `LLM_chat_html` | Main chat form. Displays messages via an embedded WebBrowser, handles agent events (`OnAnswer`, `OnReasoning`, `OnToolCalling`, `OnThinking`, `OnTick`). |
| `LLM_config` | LLM settings form (URL, model, API key, system prompt). |
| `m_toolbar` / `c_buttons` | Creating the "LLM" toolbar in Visio and handling button clicks. |
| `m_json` | JSON utilities based on `JsonConverter` from VBA-JSON. |
| `m_FileIO` | File I/O in UTF-8, path resolution, folder creation. |
| `m_Constants` | Global constants (debug mode, registry keys, default values). |
| `m_Utils` | Logging, reading/writing registry settings. |
| `m_TextTransform` | Adding the Text Transform section to Visio shapes. |
| `t_Shapes` | Utilities for working with Visio shape cells (read, write, check existence). |
| `m_DiceRoll` | Random number generator (auxiliary module). |

### Agent Loop

1. The user enters a prompt in the chat.
2. The prompt is added to the message history with the `user` role.
3. The history (including system prompt and tool descriptions) is sent to the LLM API.
4. If the LLM returns `tool_calls` — each call is executed by the corresponding tool, the result is added to the history with the `tool` role, and the loop repeats.
5. If the LLM returns a text response without `tool_calls` — this is the final answer, the loop ends.
6. A turn limit (30 by default) prevents infinite loops.

---

## Agent Tools

The agent has access to the following tools for manipulating the Visio document:

| Tool name | Description |
|---|---|
| `get_page_size` | Returns the width and height of the active page in inches. |
| `get_shapes_list` | Returns a list of all shapes on the active page with their IDs and type (polygon / line). |
| `draw_poly_line` | Draws a polyline (or a polygon if the first and last points match). Parameters: point coordinates, line weight, line color, fill color. |
| `draw_text_box` | Creates a text shape (label) by rectangle coordinates. Parameters: corner coordinates, text, rotation angle. |
| `set_format_shape` | Batch-sets values for arbitrary Visio shape cells. Format: `CELL1:value1;CELL2:value2;...` |
| `read_shape_cells` | Reads values of specified shape cells. Returns `CellName:Value` pairs. |
| `read_shape_text` | Reads the text (caption) of a specified shape. |
| `set_shape_text` | Sets the text (caption) for a specified shape. |
| `read_text_file` | Reads a text file (UTF-8) by absolute or relative path. |
| `write_text_file` | Writes text to a file (UTF-8). Creates missing folders, overwrites existing files. |

---

## Installation

### Requirements

- Microsoft Visio (2007 or newer)
- Access to an OpenAI-compatible API (e.g., [aitunnel.ru](https://aitunnel.ru), OpenAI, DeepSeek, etc.)

### Steps

1. Copy the `visio_smart_redactor.vss` file to your Visio stencils folder (usually `Documents\My Shapes\`).
2. Open Visio and drag a shape from the stencil onto the document page — this activates the project code.
3. An **"LLM"** toolbar will appear in Visio with two buttons:
   - **Chat** — open the agent chat window.
   - **Settings** — open LLM settings.

> When the document is closed, the toolbar is automatically removed.

---

## Configuration

Click the **Settings** button on the "LLM" toolbar, or open the `LLM_config` form from the VBA editor.

![](imgs/llm_settings_form.jpg)

### Parameters

| Parameter | Description | Default value |
|---|---|---|
| **API URL** | OpenAI-compatible API endpoint URL | `https://api.aitunnel.ru/v1/chat/completions` |
| **Model (main)** | Identifier of the main ("smart") model | `deepseek-v4-pro` |
| **Model (fast)** | Identifier of the fast model | `deepseek-v4-flash` |
| **API Key** | API authorization key | *(empty)* |
| **System Prompt** | System prompt defining agent behavior | *(empty)* |

Settings are stored in the Windows registry at:

```
HKEY_CURRENT_USER\Software\LLMVisio\
```

> **Important:** The API key is stored in the registry in plain text. Make sure you have sufficient system access rights.

---

## Usage

1. **Open the chat** — click the **Chat** button on the "LLM" toolbar.

![](imgs/llm_toolbar.jpg)

2. **Select a model** — a toggle in the chat window lets you choose between the main and fast model.
3. **Enter a prompt** — for example:
   - "Draw a 5×3 cm rectangle in the center of the page"
   - "Show the list of all shapes on the page"
   - "Change the text of shape with ID 3 to 'Hello, world!'"
   - "Read the file notes.txt and create a diagram from it"

![](imgs/llm_chat.jpg)

4. **Add a file to context** — the add-file button lets you load a text file (`.txt`, `.md`, `.qmd`) into the LLM context (RAG). The file content is added to the message history with the `system` role.
5. **Clear context** — the clear button resets the message history and reinitializes the agent.

### Chat Visualization

- **User messages** — green bubbles on the right.
- **Agent responses** — white bubbles on the left.
- **Reasoning** — collapsible `[Reasoning]` block.
- **Tool results** — collapsible `[Tool]` block.
- **Work indicator** — a spinning character in the window title while waiting for the LLM response.

---

## Project Structure

```
visio_redactor_agent/
├── visio_smart_redactor.vss      # Visio stencil file (contains all project code)
├── LICENSE                       # MIT License
├── README.md                     # README (English)
├── README.ru.md                  # README (Russian)
└── code/                         # VBA source code (exported modules)
    ├── ThisDocument.bas          # Document open/close handlers
    ├── m_Constants.bas           # Global constants
    ├── m_toolbar.bas             # Toolbar creation/removal
    ├── c_buttons.cls             # Toolbar button click handling
    ├── clsHarness.cls            # Main agent class (loop, HTTP, parsing)
    ├── clsHarnessToolsHandler.cls # Tool registry and dispatcher
    ├── clsToolBase.cls           # Abstract base class for tools
    ├── clsTool_PageSize.cls      # Tool: page size
    ├── clsTool_ShapesList.cls    # Tool: shapes list
    ├── clsTool_DrawPolyLine.cls  # Tool: draw polyline/polygon
    ├── clsTool_DrawTextBox.cls   # Tool: create text label
    ├── clsTool_FormatShape.cls   # Tool: format shape cells
    ├── clsTool_ReadShapeCells.cls # Tool: read shape cells
    ├── clsTool_ReadShapeText.cls # Tool: read shape text
    ├── clsTool_SetShapeText.cls  # Tool: set shape text
    ├── clsTool_ReadTextFile.cls  # Tool: read text file
    ├── clsTool_WriteTextFile.cls # Tool: write text file
    ├── LLM_chat_html.frm/.frx    # Chat form
    ├── LLM_config.frm/.frx       # Settings form
    ├── m_json.bas                # JSON utilities (based on VBA-JSON)
    ├── m_FileIO.bas              # File operations (UTF-8)
    ├── m_Utils.bas               # Utilities (logging, registry)
    ├── m_TextTransform.bas       # Text Transform section for shapes
    ├── m_DiceRoll.bas            # Random number generator
    ├── t_Shapes.bas              # Visio cell utilities
    └── JsonConverter.bas         # VBA-JSON module (third-party library)
```

---

## Example Results

![](imgs/example1.jpg)

**LLM**: GPT 5.6 sol. **Prompt**: *Draw a floor plan of a house consisting of two rooms and an extension. Draw doors and windows, indicate the main entrance. Format as a blueprint.*

![](imgs/example2.jpg)

**LLM**: GPT 5.6 sol. **Prompt**: *Draw a seascape: the horizon, the sea, fish swimming in the sea, the sun and clouds above the sea. Write a description of the drawing.*

---

## Developing Custom Tools

The project uses the **Tool** pattern based on the abstract class `clsToolBase`. Each tool is a separate class module implementing the `clsToolBase` interface. To add a new tool, follow three steps.

### Step 1. Create the tool class

Create a new class module (e.g., `clsTool_MyTool.cls`) and add `Implements clsToolBase`. You need to override four members:

| Member | Type | Purpose |
|---|---|---|
| `clsToolBase_Name` | `Property Get` | Tool name that the LLM sees in `tools[]`. Must be unique. |
| `clsToolBase_Description` | `Property Get` | Description for the LLM: when and why to call this tool. The quality of the description determines how correctly the agent will choose the tool. |
| `clsToolBase_Schema` | `Property Get` | JSON Schema of parameters as a string. Defines which arguments the LLM will pass to the tool. |
| `clsToolBase_Execute` | `Function` | Main logic. Takes a JSON string with arguments, returns a JSON string with the result. |

### Step 2. Implement the logic

```vb
Option Explicit
Implements clsToolBase

' Tool name — what the LLM sees in tools[]
Private Property Get clsToolBase_Name() As String
    clsToolBase_Name = "my_tool"
End Property

' Description for the LLM: when and why to call
Private Property Get clsToolBase_Description() As String
    clsToolBase_Description = "Does something useful. " & _
        "Call this tool when the user asks ..."
End Property

' JSON Schema of parameters
Private Property Get clsToolBase_Schema() As String
    clsToolBase_Schema = "{" & _
        """type"":""object""," & _
        """properties"":{" & _
            """param1"":{""type"":""string"",""description"":""Parameter description""}" & _
        "}," & _
        """required"": [""param1""]" & _
    "}"
End Property

' Tool execution
Private Function clsToolBase_Execute(ByVal argsJson As String) As String
    On Error GoTo ErrHandler

    ' Extract arguments from JSON
    Dim param1 As String
    param1 = ExtractJsonString(argsJson, "param1")

    ' ... tool logic ...

    ' Return result as JSON
    clsToolBase_Execute = "{""status"":""ok"",""message"":""Done""}"
    Exit Function

ErrHandler:
    clsToolBase_Execute = "{""status"":""error"",""message"":""" & Err.Description & """}"
End Function
```

### Recommendations

- **Extracting arguments.** Use the `ExtractJsonString(json, key)` function from the `m_json` module — it recursively searches for a key in JSON and returns the value as a string.
- **Escaping strings in results.** If the returned message contains quotes, newlines, or backslashes, use `JsonString(s)` from the `m_json` module for escaping. For typical responses, you can use the ready-made `MakeOk(msg)` and `MakeError(msg)` functions.
- **Result format.** Return a JSON object with `status` (`"ok"` or `"error"`) and `message` fields — the agent and LLM expect this format.
- **Error handling.** Always use `On Error GoTo ErrHandler` so that on failure the tool returns a clear error message instead of crashing the agent loop.
- **Description for the LLM.** Write the description so the LLM understands when to call the tool. Specify units, data formats, and default values.
- **Auxiliary modules.** Available: `t_Shapes` (shape cell operations), `m_FileIO` (UTF-8 files), `m_TextTransform` (Text Transform section), `m_Utils` (logging, registry).

### Step 3. Register the tool

Add registration in the `Init()` method of the `clsHarnessToolsHandler` class:

```vb
Public Sub Init()
    Set tools = New Collection

    ' ... existing tools ...

    ' Your new tool
    Dim toolInstance As clsToolBase
    Set toolInstance = New clsTool_MyTool
    tools.Add toolInstance, toolInstance.name

End Sub
```

After this, the tool will automatically:
- be included in the `tools[]` JSON schema sent to the LLM;
- be available for calling by name via the `Execute` dispatcher;
- receive arguments from the LLM and return the result to the message history.

---

## Dependencies

| Dependency | Purpose | Source |
|---|---|---|
| [VBA-Harness](https://github.com/Obsidian-pb/vba_harness) | Core LLM agent architecture (clsHarness, clsHarnessToolsHandler, clsToolBase) | Included in the project |
| [VBA-JSON](https://github.com/VBA-tools/VBA-JSON) | JSON parsing and serialization (`JsonConverter`) | Included in the project (`JsonConverter.bas`) |
| MSXML2.ServerXMLHTTP | HTTP requests to LLM API | Built into Windows |
| ADODB.Stream | Reading/writing UTF-8 files | Built into Windows |
| Microsoft.Scripting.Runtime | Dictionary object | Built into Office |
| Scripting.FileSystemObject | File operations | Built into Windows |
| WebBrowser Control | HTML-based chat rendering | Built into Office |

---

## License

This project is licensed under the **MIT License**. See the [LICENSE](LICENSE) file.

Copyright © 2026 Obsidian-pb