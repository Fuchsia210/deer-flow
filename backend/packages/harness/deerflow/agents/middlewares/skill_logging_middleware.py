"""Middleware for logging skill usage and execution.

This middleware tracks:
- When skills are loaded/read
- Which skills are used in conversations
- Skill execution timeline and outcomes
- Script execution (bash/python commands) - with REAL terminal output
- All print statements and error messages from skills
"""

import json
import logging
import os
from datetime import datetime
from pathlib import Path
from typing import TYPE_CHECKING, Any, NotRequired, override

from langchain.agents import AgentState
from langchain.agents.middleware import AgentMiddleware
from langgraph.runtime import Runtime

from deerflow.config import get_app_config
from deerflow.config.paths import get_paths
from deerflow.skills.skill_log_manager import is_error_from_output

if TYPE_CHECKING:
    from deerflow.config.app_config import AppConfig

logger = logging.getLogger(__name__)

# Tools that execute commands - we capture their raw output
_EXECUTION_TOOLS = ["bash", "python"]


class SkillLoggingState(AgentState):
    """State schema for skill logging middleware."""

    skill_logs: NotRequired[list[dict[str, Any]] | None]
    active_skill: NotRequired[str | None]


class SkillLoggingMiddleware(AgentMiddleware[SkillLoggingState]):
    """Middleware to log skill usage and execution events."""

    state_schema = SkillLoggingState

    def __init__(self, *, app_config: "AppConfig | None" = None):
        super().__init__()
        self._app_config = app_config
        self._log_dir: Path | None = None
        self._ensure_log_directory()

    def _ensure_log_directory(self):
        """Ensure the skill log directory exists."""
        try:
            if self._app_config is None:
                config = get_app_config()
            else:
                config = self._app_config
            
            # Get or create log directory
            paths = get_paths()
            skill_log_dir = paths.base_dir / "logs" / "skills"
            skill_log_dir.mkdir(parents=True, exist_ok=True)
            self._log_dir = skill_log_dir
        except Exception as e:
            logger.debug(f"Failed to create skill log directory: {e}")

    def _get_thread_id(self, runtime: Runtime) -> str:
        """Extract thread ID from runtime context."""
        thread_id = None
        if runtime.context:
            thread_id = runtime.context.get("thread_id")
        if not thread_id and runtime.config:
            thread_id = runtime.config.get("configurable", {}).get("thread_id")
        return thread_id or "unknown_thread"

    def _get_skill_name_from_path(self, path: str) -> str | None:
        """Extract skill name from a file path.
        
        Skill paths typically look like: /mnt/skills/public/skill-name/SKILL.md
        """
        try:
            parts = path.split("/")
            # Find the SKILL.md file and get its parent directory name
            for i, part in enumerate(parts):
                if part == "SKILL.md" and i >= 2:
                    return parts[i-1]
        except Exception:
            pass
        return None

    def _is_skill_file(self, path: str) -> bool:
        """Check if the path points to a skill-related file."""
        if not path:
            return False
        path_lower = path.lower()
        return ("/mnt/skills" in path_lower or "skills/" in path_lower) and (
            "skill.md" in path_lower or 
            "/references/" in path_lower or 
            "/templates/" in path_lower
        )

    def _log_skill_event(
        self,
        event_type: str,
        skill_name: str | None,
        thread_id: str,
        state: SkillLoggingState,
        extra: dict[str, Any] | None = None
    ):
        """Log a skill event to both state and file."""
        timestamp = datetime.utcnow().isoformat()
        event = {
            "timestamp": timestamp,
            "event_type": event_type,
            "skill_name": skill_name,
            "thread_id": thread_id,
            **(extra or {})
        }

        # Update state
        logs = state.get("skill_logs") or []
        logs.append(event)
        state["skill_logs"] = logs

        # Update active skill
        if event_type == "skill_loaded":
            state["active_skill"] = skill_name

        # Log to file if possible
        if self._log_dir:
            self._write_log_to_file(thread_id, event)

        logger.debug(f"Skill event: {event_type} - {skill_name} (thread: {thread_id})")

    def _write_log_to_file(self, thread_id: str, event: dict[str, Any]):
        """Write a log event to a JSONL file."""
        if not self._log_dir:
            return

        try:
            log_file = self._log_dir / f"skill_logs_{datetime.now().strftime('%Y%m%d')}.jsonl"
            with open(log_file, "a", encoding="utf-8") as f:
                f.write(json.dumps(event, ensure_ascii=False) + "\n")
        except Exception as e:
            logger.debug(f"Failed to write skill log to file: {e}")

    def _analyze_tool_calls(self, state: SkillLoggingState) -> list[dict[str, Any]]:
        """Analyze messages for tool calls related to skills."""
        events = []
        messages = state.get("messages", [])
        
        for msg in messages:
            if hasattr(msg, "tool_calls") and msg.tool_calls:
                for tc in msg.tool_calls:
                    tool_name = tc.get("name")
                    args = tc.get("args", {})
                    
                    if tool_name == "read_file":
                        path = args.get("path", "")
                        if self._is_skill_file(path):
                            skill_name = self._get_skill_name_from_path(path)
                            if skill_name:
                                events.append({
                                    "type": "skill_file_read",
                                    "skill_name": skill_name,
                                    "file_path": path
                                })
        
        return events

    def _analyze_tool_outputs(self, state: SkillLoggingState, active_skill: str | None) -> list[dict[str, Any]]:
        """Analyze tool outputs - capturing REAL terminal output (stdout + stderr + prints)."""
        events = []
        messages = state.get("messages", [])
        
        for i, msg in enumerate(messages):
            # Check if this is a tool response (tool_message)
            if hasattr(msg, "type") and msg.type == "tool":
                # Find the corresponding tool call
                tool_call_id = getattr(msg, "tool_call_id", None)
                tool_name = getattr(msg, "name", None)
                raw_output = getattr(msg, "content", "")  # This is the REAL terminal output!
                
                if tool_name in _EXECUTION_TOOLS:
                    # Find the tool_call message
                    command = ""
                    description = ""
                    if i > 0:
                        prev_msg = messages[i - 1]
                        if hasattr(prev_msg, "tool_calls") and prev_msg.tool_calls:
                            for tc in prev_msg.tool_calls:
                                if tc.get("id") == tool_call_id:
                                    args = tc.get("args", {})
                                    command = args.get("command", "")
                                    description = args.get("description", "")
                                    break
                    
                    # Check if command is skill-related or we have active skill
                    is_skill_related = active_skill is not None or self._is_skill_file(command)
                    
                    if is_skill_related:
                        # Detect if there was an error - check log levels and error patterns
                        is_error_output = is_error_from_output(raw_output)
                        
                        events.append({
                            "type": "command_executed",
                            "skill_name": active_skill,
                            "tool_name": tool_name,
                            "description": description,
                            "command": command[:1000] if command else "",  # Save more of the command
                            "raw_output": raw_output,  # SAVE THE FULL RAW TERMINAL OUTPUT!
                            "is_error": is_error_output,
                        })
        
        return events

    @override
    def before_model(self, state: SkillLoggingState, runtime: Runtime) -> dict | None:
        """Hook before model invocation to track skill usage."""
        try:
            thread_id = self._get_thread_id(runtime)
            active_skill = state.get("active_skill")
            
            # Check for new skill-related tool calls
            tool_events = self._analyze_tool_calls(state)
            for event in tool_events:
                # Only log if we haven't logged this skill load yet
                logs = state.get("skill_logs") or []
                already_logged = any(
                    log.get("event_type") == "skill_loaded" and 
                    log.get("skill_name") == event["skill_name"]
                    for log in logs
                )
                
                if not already_logged:
                    self._log_skill_event(
                        "skill_loaded",
                        event["skill_name"],
                        thread_id,
                        state,
                        {"file_path": event["file_path"]}
                    )
            
            # Check for new command executions - with RAW terminal output
            if active_skill:
                command_events = self._analyze_tool_outputs(state, active_skill)
                logs = state.get("skill_logs") or []
                
                # Track which command events we've already logged
                logged_command_key = set()
                for log in logs:
                    if log.get("event_type") == "command_executed":
                        key = (log.get("command", ""), log.get("tool_name", ""))
                        logged_command_key.add(key)
                
                for event in command_events:
                    key = (event["command"], event["tool_name"])
                    if key not in logged_command_key:
                        self._log_skill_event(
                            "command_executed",
                            event["skill_name"],
                            thread_id,
                            state,
                            {
                                "tool_name": event["tool_name"],
                                "description": event["description"],
                                "command": event["command"],
                                "raw_output": event["raw_output"],  # RAW TERMINAL OUTPUT!
                                "is_error": event["is_error"],
                            }
                        )
                        logged_command_key.add(key)
            
            return None
        except Exception as e:
            logger.debug(f"Error in skill logging before_model: {e}")
            return None

    @override
    async def abefore_model(self, state: SkillLoggingState, runtime: Runtime) -> dict | None:
        """Async version of before_model."""
        return self.before_model(state, runtime)

    @override
    def after_model(self, state: SkillLoggingState, runtime: Runtime) -> dict | None:
        """Hook after model invocation to log skill execution outcomes."""
        try:
            thread_id = self._get_thread_id(runtime)
            active_skill = state.get("active_skill")
            
            if active_skill:
                # Check for command executions that may have happened since last hook
                command_events = self._analyze_tool_outputs(state, active_skill)
                logs = state.get("skill_logs") or []
                
                # Track which command events we've already logged
                logged_command_key = set()
                for log in logs:
                    if log.get("event_type") == "command_executed":
                        key = (log.get("command", ""), log.get("tool_name", ""))
                        logged_command_key.add(key)
                
                for event in command_events:
                    key = (event["command"], event["tool_name"])
                    if key not in logged_command_key:
                        self._log_skill_event(
                            "command_executed",
                            event["skill_name"],
                            thread_id,
                            state,
                            {
                                "tool_name": event["tool_name"],
                                "description": event["description"],
                                "command": event["command"],
                                "raw_output": event["raw_output"],
                                "is_error": event["is_error"],
                            }
                        )
                        logged_command_key.add(key)
            
            # Check if this is a final response (no tool calls)
            messages = state.get("messages", [])
            last_msg = messages[-1] if messages else None
            
            if (last_msg and hasattr(last_msg, "type") and 
                last_msg.type == "ai" and not getattr(last_msg, "tool_calls", [])):
                
                if active_skill:
                    # Check if there were any errors during execution
                    logs = state.get("skill_logs") or []
                    had_error = any(
                        log.get("event_type") == "command_executed" and log.get("is_error")
                        for log in logs
                    )
                    
                    # Log skill completion
                    self._log_skill_event(
                        "skill_completed",
                        active_skill,
                        thread_id,
                        state,
                        {
                            "message_count": len(messages),
                            "had_error": had_error,
                        }
                    )
                    # Clear active skill
                    state["active_skill"] = None
            
            return None
        except Exception as e:
            logger.debug(f"Error in skill logging after_model: {e}")
            return None

    @override
    async def aafter_model(self, state: SkillLoggingState, runtime: Runtime) -> dict | None:
        """Async version of after_model."""
        return self.after_model(state, runtime)
