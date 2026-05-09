"""Skill Log Manager - Utilities for analyzing skill execution logs."""

import json
import logging
import re
from collections import defaultdict
from datetime import datetime, timedelta
from pathlib import Path
from typing import Any

from deerflow.config import get_app_config
from deerflow.config.paths import get_paths

logger = logging.getLogger(__name__)

# Log level patterns to match
LOG_LEVEL_PATTERNS = [
    (logging.CRITICAL, r"\b(?:CRITICAL|FATAL)\b", re.IGNORECASE),
    (logging.ERROR, r"\b(?:ERROR)\b", re.IGNORECASE),
    (logging.WARNING, r"\b(?:WARNING|WARN)\b", re.IGNORECASE),
    (logging.INFO, r"\b(?:INFO)\b", re.IGNORECASE),
    (logging.DEBUG, r"\b(?:DEBUG)\b", re.IGNORECASE),
]


def parse_log_level_from_output(output: str) -> int:
    """Parse log level from terminal output.
    
    Args:
        output: The raw terminal output to parse.
        
    Returns:
        The logging level (e.g., logging.INFO, logging.ERROR), 
        defaults to logging.INFO if no level is detected.
    """
    if not output:
        return logging.INFO
    
    # Check each line for log level patterns
    lines = output.split("\n")
    for line in lines:
        line = line.strip()
        if not line:
            continue
        for level, pattern, flags in LOG_LEVEL_PATTERNS:
            if re.search(pattern, line, flags=flags):
                return level
    
    # If no log level detected, default to INFO
    return logging.INFO


def is_error_from_output(output: str) -> bool:
    """Check if output contains error level logs.
    
    Args:
        output: The raw terminal output to check.
        
    Returns:
        True if output contains errors, False otherwise.
    """
    if not output:
        return False
    
    # Check if log level is ERROR or higher
    level = parse_log_level_from_output(output)
    return level >= logging.ERROR


class SkillLogManager:
    """Manager for skill execution logs."""

    def __init__(self, log_dir: Path | None = None):
        self.config = get_app_config().skill_logging
        if log_dir is None:
            if self.config.log_dir:
                log_dir = Path(self.config.log_dir)
            else:
                paths = get_paths()
                log_dir = paths.base_dir / "logs" / "skills"
        self.log_dir = Path(log_dir)

    def _get_log_files(self, days: int = 30) -> list[Path]:
        """Get all log files within the specified number of days."""
        if not self.log_dir.exists():
            return []
        
        cutoff = datetime.now() - timedelta(days=days)
        log_files = []
        
        for file in self.log_dir.glob("skill_logs_*.jsonl"):
            try:
                date_str = file.stem.split("_")[-1]
                file_date = datetime.strptime(date_str, "%Y%m%d")
                if file_date >= cutoff:
                    log_files.append(file)
            except Exception as e:
                logger.debug(f"Skipping invalid log file {file}: {e}")
        
        return sorted(log_files)

    def load_logs(self, days: int = 30, thread_id: str | None = None, skill_name: str | None = None) -> list[dict[str, Any]]:
        """Load skill logs with optional filtering."""
        logs = []
        for log_file in self._get_log_files(days):
            try:
                with open(log_file, encoding="utf-8") as f:
                    for line in f:
                        line = line.strip()
                        if not line:
                            continue
                        try:
                            log_entry = json.loads(line)
                            
                            # Apply filters
                            if thread_id and log_entry.get("thread_id") != thread_id:
                                continue
                            if skill_name and log_entry.get("skill_name") != skill_name:
                                continue
                            
                            logs.append(log_entry)
                        except json.JSONDecodeError as e:
                            logger.debug(f"Invalid JSON in {log_file}: {e}")
            except Exception as e:
                logger.error(f"Error reading log file {log_file}: {e}")
        
        return sorted(logs, key=lambda x: x.get("timestamp", ""))

    def get_skill_usage_stats(self, days: int = 30) -> dict[str, Any]:
        """Get statistics about skill usage."""
        logs = self.load_logs(days)
        
        stats = {
            "total_events": len(logs),
            "skills_loaded": 0,
            "skills_completed": 0,
            "commands_executed": 0,
            "errors_count": 0,
            "skill_counts": defaultdict(int),
            "skill_error_counts": defaultdict(int),
            "skill_command_counts": defaultdict(int),
            "thread_counts": set(),
            # Log level statistics
            "log_level_counts": defaultdict(int),
            "skill_log_level_counts": defaultdict(lambda: defaultdict(int)),
        }
        
        for log in logs:
            event_type = log.get("event_type")
            skill_name = log.get("skill_name")
            thread_id = log.get("thread_id")
            
            if thread_id:
                stats["thread_counts"].add(thread_id)
            
            if event_type == "skill_loaded":
                stats["skills_loaded"] += 1
                if skill_name:
                    stats["skill_counts"][skill_name] += 1
            elif event_type == "skill_completed":
                stats["skills_completed"] += 1
                if log.get("had_error") and skill_name:
                    stats["skill_error_counts"][skill_name] += 1
            elif event_type == "command_executed":
                stats["commands_executed"] += 1
                if skill_name:
                    stats["skill_command_counts"][skill_name] += 1
                # Check for error - support both new and old field names
                if log.get("is_error") or log.get("has_error"):
                    stats["errors_count"] += 1
                    if skill_name:
                        stats["skill_error_counts"][skill_name] += 1
                
                # Parse log level from output and count
                raw_output = log.get("raw_output", log.get("error_summary", ""))
                log_level = parse_log_level_from_output(raw_output)
                log_level_name = logging.getLevelName(log_level)
                stats["log_level_counts"][log_level_name] += 1
                if skill_name:
                    stats["skill_log_level_counts"][skill_name][log_level_name] += 1
        
        stats["thread_count"] = len(stats["thread_counts"])
        stats["skill_counts"] = dict(sorted(stats["skill_counts"].items(), key=lambda x: -x[1]))
        stats["skill_error_counts"] = dict(sorted(stats["skill_error_counts"].items(), key=lambda x: -x[1]))
        stats["skill_command_counts"] = dict(sorted(stats["skill_command_counts"].items(), key=lambda x: -x[1]))
        stats["log_level_counts"] = dict(stats["log_level_counts"])
        # Convert nested defaultdicts to regular dicts
        stats["skill_log_level_counts"] = {
            skill: dict(counts) 
            for skill, counts in stats["skill_log_level_counts"].items()
        }
        
        return stats

    def get_thread_skill_history(self, thread_id: str) -> list[dict[str, Any]]:
        """Get skill usage history for a specific thread."""
        return self.load_logs(thread_id=thread_id)

    def get_command_outputs(self, thread_id: str | None = None, skill_name: str | None = None) -> list[dict[str, Any]]:
        """Get all command execution outputs with raw terminal content."""
        logs = self.load_logs(thread_id=thread_id)
        command_logs = [log for log in logs if log.get("event_type") == "command_executed"]
        if skill_name:
            command_logs = [log for log in command_logs if log.get("skill_name") == skill_name]
        return command_logs

    def print_command_output(self, log: dict[str, Any]) -> None:
        """Print a single command execution with raw output."""
        print("\n" + "-" * 80)
        print(f"TIMESTAMP: {log.get('timestamp', 'N/A')}")
        print(f"SKILL: {log.get('skill_name', 'N/A')}")
        print(f"TOOL: {log.get('tool_name', 'N/A')}")
        print(f"DESCRIPTION: {log.get('description', 'N/A')}")
        print(f"COMMAND: {log.get('command', 'N/A')}")
        print(f"ERROR: {'YES' if log.get('is_error') or log.get('has_error') else 'NO'}")
        print("\n--- RAW TERMINAL OUTPUT ---")
        print(log.get("raw_output", log.get("error_summary", "(No output)")))
        print("-" * 80 + "\n")

    def cleanup_old_logs(self, days: int | None = None) -> int:
        """Remove log files older than the specified number of days."""
        if days is None:
            days = self.config.max_days_to_keep
        
        cutoff = datetime.now() - timedelta(days=days)
        deleted_count = 0
        
        if not self.log_dir.exists():
            return 0
        
        for file in self.log_dir.glob("skill_logs_*.jsonl"):
            try:
                date_str = file.stem.split("_")[-1]
                file_date = datetime.strptime(date_str, "%Y%m%d")
                if file_date < cutoff:
                    file.unlink()
                    deleted_count += 1
                    logger.info(f"Deleted old log file: {file}")
            except Exception as e:
                logger.warning(f"Could not process file {file} for cleanup: {e}")
        
        return deleted_count

    def print_summary(self, days: int = 30) -> None:
        """Print a summary of skill usage."""
        stats = self.get_skill_usage_stats(days)
        
        print("\n" + "=" * 70)
        print("SKILL USAGE SUMMARY")
        print("=" * 70)
        print(f"Total events: {stats['total_events']}")
        print(f"Skills loaded: {stats['skills_loaded']}")
        print(f"Skills completed: {stats['skills_completed']}")
        print(f"Commands executed: {stats['commands_executed']}")
        print(f"Errors encountered: {stats['errors_count']}")
        print(f"Unique threads: {stats['thread_count']}")
        
        # Log level summary
        if stats["log_level_counts"]:
            print("\nLog level distribution:")
            for level_name in ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]:
                count = stats["log_level_counts"].get(level_name, 0)
                if count > 0:
                    print(f"  - {level_name}: {count}")
        
        print("\nNote: All command outputs are captured RAW from terminal!")
        print("      Use get_command_outputs() or print_command_output() to view raw output.")
        
        if stats["skill_counts"]:
            print("\nSkill usage by name:")
            for skill, count in stats["skill_counts"].items():
                cmd_count = stats["skill_command_counts"].get(skill, 0)
                err_count = stats["skill_error_counts"].get(skill, 0)
                # Log level summary per skill
                skill_log_levels = stats["skill_log_level_counts"].get(skill, {})
                log_level_str = ", ".join([f"{level}: {count}" for level, count in skill_log_levels.items()])
                print(f"  - {skill}: {count} loads, {cmd_count} commands, {err_count} errors")
                if log_level_str:
                    print(f"    Log levels: {log_level_str}")
        
        if stats["skill_error_counts"]:
            print("\nSkills with errors:")
            for skill, count in stats["skill_error_counts"].items():
                print(f"  - {skill}: {count} errors")
        
        print("=" * 70 + "\n")


def get_skill_log_manager() -> SkillLogManager:
    """Get the skill log manager instance."""
    return SkillLogManager()
