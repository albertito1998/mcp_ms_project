# MCP Server for Microsoft Project

An MCP (Model Context Protocol) server that allows Claude and other LLM clients to read Microsoft Project files (`.mpp`) directly, without needing MS Project open.

Built on top of [mpxj](https://mpxj.org/) — a battle-tested Java library with Python bindings that supports reading `.mpp`, `.mpt`, `.mpx`, `.xml`, `.xer` and other project file formats.

## Features

| Tool | Description |
|------|-------------|
| `get_project_summary` | Project title, author, dates, task and resource counts |
| `get_tasks` | All tasks with WBS, dates, % complete, duration, cost, predecessors |
| `get_resources` | People, equipment and materials with rates and units |
| `get_assignments` | Resource-to-task assignments with work and cost |
| `get_critical_path` | Tasks on the critical path |
| `get_overdue_tasks` | Tasks past their finish date and not 100% complete |

## Requirements

- Python 3.10+
- Java 11+ (required by mpxj via JPype)
- pip packages: `mcp`, `mpxj`, `JPype1`

## Installation

```bash
# 1. Clone the repo
git clone https://github.com/albertito1998/mcp_ms_project.git
cd mcp_ms_project

# 2. Install dependencies
pip install -r requirements.txt
```

> Make sure Java is installed and available in your PATH (`java --version`).

## Usage with Claude Code

Add the server to Claude Code:

```bash
claude mcp add msproject -- python /path/to/msproject_server.py
```

## Usage with Claude Desktop

Add the following to your `claude_desktop_config.json`:

```json
{
  "mcpServers": {
    "msproject": {
      "command": "python",
      "args": ["/path/to/msproject_server.py"]
    }
  }
}
```

On **Windows** the config file is at:
```
%APPDATA%\Claude\claude_desktop_config.json
```

On **macOS/Linux**:
```
~/.config/Claude/claude_desktop_config.json
```

## Example prompts

Once connected, you can ask Claude:

- *"Summarise the project at C:/Projects/plan.mpp"*
- *"Which tasks are overdue in my project?"*
- *"Show me the critical path"*
- *"Who is assigned to which tasks?"*
- *"List all tasks that are less than 50% complete"*

## Supported file formats

| Extension | Format |
|-----------|--------|
| `.mpp` | Microsoft Project (all versions) |
| `.mpt` | Microsoft Project Template |
| `.mpx` | Microsoft Project Exchange |
| `.xml` | Microsoft Project XML export |
| `.xer` | Primavera XER |
| `.pp` | Asta Powerproject |

## License

MIT
