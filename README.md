# MCP Server for Microsoft Project

An MCP (Model Context Protocol) server that allows Claude and other LLM clients to **read and write** Microsoft Project files (`.mpp`) directly, without needing MS Project open.

Built on top of [mpxj](https://mpxj.org/) — a battle-tested Java library with Python bindings that supports `.mpp`, `.mpt`, `.mpx`, `.xml`, `.xer` and other project file formats.

> **Write note:** mpxj does not support writing native `.mpp` binary files. Write operations save to **MSPDI XML** (`.xml`) format, which Microsoft Project opens natively. Open the `.xml` in MS Project and use *File → Save As* to convert back to `.mpp`.

## Tools

### Read

| Tool | Description |
|------|-------------|
| `get_project_summary` | Title, author, company, dates, task and resource counts |
| `get_tasks` | All tasks with WBS, dates, % complete, duration, cost, predecessors |
| `get_resources` | People, equipment and materials with rates and units |
| `get_assignments` | Resource-to-task assignments with work and cost |
| `get_critical_path` | Tasks on the critical path |
| `get_overdue_tasks` | Tasks past their finish date and not 100% complete |

### Write

| Tool | Description |
|------|-------------|
| `add_task` | Add a new task (or subtask under a parent) |
| `update_task` | Edit name, dates, % complete, duration, notes, milestone flag |
| `delete_task` | Remove a task |
| `add_resource` | Add a person, equipment, or material resource |
| `update_resource` | Edit resource name, email, max units, group, notes |
| `delete_resource` | Remove a resource (also removes its assignments) |
| `assign_resource` | Assign a resource to a task with optional units |
| `remove_assignment` | Unassign a resource from a task |
| `update_project_properties` | Edit title, author, company, start date, status date |
| `convert_project` | Convert to `.xml`, `.mpx`, `.json`, or `.xer` |

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

```bash
claude mcp add msproject -- python /path/to/msproject_server.py
```

## Usage with Claude Desktop

Add to `claude_desktop_config.json`:

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

**Windows:** `%APPDATA%\Claude\claude_desktop_config.json`  
**macOS/Linux:** `~/.config/Claude/claude_desktop_config.json`

## Example prompts

**Reading:**
- *"Summarise the project at C:/Projects/plan.mpp"*
- *"Which tasks are overdue?"*
- *"Show me the critical path"*
- *"Who is assigned to which tasks?"*

**Writing:**
- *"Add a task called 'UAT' starting 2025-05-01 for 3 days"*
- *"Mark task 5 as 75% complete"*
- *"Add Alice as a resource and assign her to task 3"*
- *"Delete task 8 from the project"*
- *"Update the project title to 'Q2 Rollout' and set status date to today"*
- *"Convert plan.mpp to plan.json"*

## Supported input formats

| Extension | Format |
|-----------|--------|
| `.mpp` | Microsoft Project (all versions) |
| `.mpt` | Microsoft Project Template |
| `.mpx` | Microsoft Project Exchange |
| `.xml` | Microsoft Project XML (MSPDI) |
| `.xer` | Primavera XER |
| `.pp` | Asta Powerproject |

## Supported output formats

| Extension | Format |
|-----------|--------|
| `.xml` | Microsoft Project XML (MSPDI) — opens in MS Project |
| `.mpx` | Microsoft Project Exchange (legacy) |
| `.json` | JSON |
| `.xer` | Primavera XER |

## License

MIT
