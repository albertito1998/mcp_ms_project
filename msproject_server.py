"""
MCP Server for Microsoft Project (.mpp) files using mpxj.
Provides tools to read tasks, resources, assignments, and project properties.
"""

import sys
import json
from pathlib import Path
from datetime import datetime
from mcp.server.fastmcp import FastMCP

mcp = FastMCP("msproject")


def _format_date(date_obj):
    if date_obj is None:
        return None
    try:
        if hasattr(date_obj, 'isoformat'):
            return date_obj.isoformat()
        return str(date_obj)
    except Exception:
        return str(date_obj)


def _format_duration(duration_obj):
    if duration_obj is None:
        return None
    try:
        return f"{duration_obj.getDuration()} {duration_obj.getUnits()}"
    except Exception:
        return str(duration_obj)


def _load_project(file_path: str):
    from mpxj import ProjectReader
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {file_path}")
    if path.suffix.lower() not in ('.mpp', '.mpt', '.mpx', '.xml', '.xer', '.pp'):
        raise ValueError(f"Unsupported file type: {path.suffix}")
    reader = ProjectReader()
    return reader.read(str(path))


@mcp.tool()
def get_project_summary(file_path: str) -> str:
    """
    Get a summary of an MS Project file: title, dates, task count, resource count.

    Args:
        file_path: Full path to the .mpp file (e.g. C:/Users/user/Documents/project.mpp)
    """
    try:
        project = _load_project(file_path)
        props = project.getProjectProperties()

        tasks = [t for t in project.getTasks() if t.getName() is not None]
        resources = [r for r in project.getResources() if r.getName() is not None]

        summary = {
            "title": str(props.getProjectTitle()) if props.getProjectTitle() else Path(file_path).stem,
            "author": str(props.getAuthor()) if props.getAuthor() else None,
            "company": str(props.getCompany()) if props.getCompany() else None,
            "start_date": _format_date(props.getStartDate()),
            "finish_date": _format_date(props.getFinishDate()),
            "status_date": _format_date(props.getStatusDate()),
            "task_count": len(tasks),
            "resource_count": len(resources),
            "currency_symbol": str(props.getCurrencySymbol()) if props.getCurrencySymbol() else None,
        }
        return json.dumps(summary, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_tasks(file_path: str, include_summary_tasks: bool = True, max_tasks: int = 200) -> str:
    """
    Get all tasks from an MS Project file with their key fields.

    Args:
        file_path: Full path to the .mpp file
        include_summary_tasks: Whether to include summary/parent tasks (default True)
        max_tasks: Maximum number of tasks to return (default 200)
    """
    try:
        project = _load_project(file_path)
        result = []

        for task in project.getTasks():
            if task.getName() is None:
                continue
            if not include_summary_tasks and task.getSummary():
                continue

            t = {
                "id": task.getID(),
                "unique_id": task.getUniqueID(),
                "name": str(task.getName()),
                "outline_level": task.getOutlineLevel(),
                "outline_number": str(task.getOutlineNumber()) if task.getOutlineNumber() else None,
                "is_summary": bool(task.getSummary()),
                "is_milestone": bool(task.getMilestone()),
                "percent_complete": float(task.getPercentageComplete()) if task.getPercentageComplete() is not None else None,
                "start": _format_date(task.getStart()),
                "finish": _format_date(task.getFinish()),
                "baseline_start": _format_date(task.getBaselineStart()),
                "baseline_finish": _format_date(task.getBaselineFinish()),
                "duration": _format_duration(task.getDuration()),
                "work": _format_duration(task.getWork()),
                "cost": float(task.getCost()) if task.getCost() is not None else None,
                "priority": str(task.getPriority()) if task.getPriority() else None,
                "notes": str(task.getNotes()) if task.getNotes() else None,
                "predecessors": str(task.getPredecessors()) if task.getPredecessors() else None,
                "wbs": str(task.getWBS()) if task.getWBS() else None,
            }
            result.append(t)
            if len(result) >= max_tasks:
                break

        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_resources(file_path: str) -> str:
    """
    Get all resources (people, equipment, materials) from an MS Project file.

    Args:
        file_path: Full path to the .mpp file
    """
    try:
        project = _load_project(file_path)
        result = []

        for resource in project.getResources():
            if resource.getName() is None:
                continue
            r = {
                "id": resource.getID(),
                "unique_id": resource.getUniqueID(),
                "name": str(resource.getName()),
                "type": str(resource.getType()) if resource.getType() else None,
                "email": str(resource.getEmailAddress()) if resource.getEmailAddress() else None,
                "max_units": float(resource.getMaxUnits()) if resource.getMaxUnits() is not None else None,
                "standard_rate": str(resource.getStandardRate()) if resource.getStandardRate() else None,
                "overtime_rate": str(resource.getOvertimeRate()) if resource.getOvertimeRate() else None,
                "group": str(resource.getGroup()) if resource.getGroup() else None,
                "notes": str(resource.getNotes()) if resource.getNotes() else None,
            }
            result.append(r)

        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_assignments(file_path: str) -> str:
    """
    Get all resource assignments (who is assigned to which tasks) from an MS Project file.

    Args:
        file_path: Full path to the .mpp file
    """
    try:
        project = _load_project(file_path)
        result = []

        for assignment in project.getResourceAssignments():
            task = assignment.getTask()
            resource = assignment.getResource()
            a = {
                "task_id": task.getID() if task else None,
                "task_name": str(task.getName()) if task else None,
                "resource_id": resource.getID() if resource else None,
                "resource_name": str(resource.getName()) if resource else None,
                "units": float(assignment.getUnits()) if assignment.getUnits() is not None else None,
                "work": _format_duration(assignment.getWork()),
                "actual_work": _format_duration(assignment.getActualWork()),
                "start": _format_date(assignment.getStart()),
                "finish": _format_date(assignment.getFinish()),
                "cost": float(assignment.getCost()) if assignment.getCost() is not None else None,
            }
            result.append(a)

        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_critical_path(file_path: str) -> str:
    """
    Get all tasks on the critical path of an MS Project file.

    Args:
        file_path: Full path to the .mpp file
    """
    try:
        project = _load_project(file_path)
        result = []

        for task in project.getTasks():
            if task.getName() is None:
                continue
            if task.getCritical():
                t = {
                    "id": task.getID(),
                    "name": str(task.getName()),
                    "start": _format_date(task.getStart()),
                    "finish": _format_date(task.getFinish()),
                    "duration": _format_duration(task.getDuration()),
                    "percent_complete": float(task.getPercentageComplete()) if task.getPercentageComplete() is not None else None,
                    "is_milestone": bool(task.getMilestone()),
                }
                result.append(t)

        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


@mcp.tool()
def get_overdue_tasks(file_path: str, status_date: str = None) -> str:
    """
    Get tasks that are behind schedule (% complete < expected based on dates).

    Args:
        file_path: Full path to the .mpp file
        status_date: Reference date in ISO format (e.g. 2025-03-28). Defaults to today.
    """
    try:
        project = _load_project(file_path)

        if status_date:
            ref_date = datetime.fromisoformat(status_date)
        else:
            ref_date = datetime.now()

        result = []
        for task in project.getTasks():
            if task.getName() is None or task.getSummary():
                continue

            finish = task.getFinish()
            if finish is None:
                continue

            try:
                finish_py = datetime(finish.getYear(), finish.getMonthValue(), finish.getDayOfMonth())
            except Exception:
                continue

            pct = float(task.getPercentageComplete()) if task.getPercentageComplete() is not None else 0.0

            if finish_py < ref_date and pct < 100.0:
                result.append({
                    "id": task.getID(),
                    "name": str(task.getName()),
                    "finish": _format_date(task.getFinish()),
                    "percent_complete": pct,
                    "is_milestone": bool(task.getMilestone()),
                })

        return json.dumps(result, indent=2, ensure_ascii=False)
    except Exception as e:
        return json.dumps({"error": str(e)})


if __name__ == "__main__":
    mcp.run()
