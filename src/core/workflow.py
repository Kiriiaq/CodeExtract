"""
Workflow Manager - Professional workflow execution system.
Handles step-by-step execution with logging, error handling, and status tracking.
"""

from dataclasses import dataclass, field
from enum import Enum, auto
from typing import Callable, Dict, List, Optional, Any
from datetime import datetime
from pathlib import Path
import traceback
import threading
from concurrent.futures import ThreadPoolExecutor
import queue


class StepStatus(Enum):
    """Status of a workflow step."""
    PENDING = auto()
    IN_PROGRESS = auto()
    COMPLETED = auto()
    SKIPPED = auto()
    ERROR = auto()
    DISABLED = auto()


@dataclass
class StepResult:
    """Result of a workflow step execution."""
    success: bool
    message: str = ""
    data: Any = None
    error: Optional[Exception] = None
    duration: float = 0.0


@dataclass
class WorkflowStep:
    """Definition of a single workflow step."""
    id: str
    name: str
    description: str
    function: Callable[..., StepResult]
    enabled: bool = True
    continue_on_error: bool = False
    timeout: Optional[float] = None
    dependencies: List[str] = field(default_factory=list)

    # Runtime state
    status: StepStatus = StepStatus.PENDING
    result: Optional[StepResult] = None
    started_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None

    def reset(self):
        """Reset step to initial state."""
        self.status = StepStatus.PENDING if self.enabled else StepStatus.DISABLED
        self.result = None
        self.started_at = None
        self.completed_at = None


class WorkflowManager:
    """
    Professional workflow manager for executing multi-step processes.
    Features: sequential execution, logging, error handling, run directories.
    """

    def __init__(self, name: str = "Workflow", output_base_dir: str = "runs"):
        self.name = name
        self.output_base_dir = Path(output_base_dir)
        self.steps: Dict[str, WorkflowStep] = {}
        self.step_order: List[str] = []

        # Execution state
        self.is_running = False
        self.should_stop = False
        self.current_step_id: Optional[str] = None
        self.run_directory: Optional[Path] = None

        # Callbacks
        self.on_step_start: Optional[Callable[[WorkflowStep], None]] = None
        self.on_step_complete: Optional[Callable[[WorkflowStep], None]] = None
        self.on_log: Optional[Callable[[str, str], None]] = None
        self.on_progress: Optional[Callable[[int, int], None]] = None

        # Configuration
        self.continue_on_error = False
        self.create_step_directories = True

        # Thread-safe log queue
        self.log_queue: queue.Queue = queue.Queue()

    def add_step(self, step: WorkflowStep) -> None:
        """Add a step to the workflow."""
        if step.id in self.steps:
            raise ValueError(f"Step with ID '{step.id}' already exists")
        self.steps[step.id] = step
        self.step_order.append(step.id)

    def remove_step(self, step_id: str) -> None:
        """Remove a step from the workflow."""
        if step_id in self.steps:
            del self.steps[step_id]
            self.step_order.remove(step_id)

    def reorder_steps(self, new_order: List[str]) -> None:
        """Reorder the steps in the workflow."""
        if set(new_order) != set(self.step_order):
            raise ValueError("New order must contain exactly the same step IDs")
        self.step_order = new_order

    def enable_step(self, step_id: str, enabled: bool = True) -> None:
        """Enable or disable a step."""
        if step_id in self.steps:
            self.steps[step_id].enabled = enabled
            self.steps[step_id].status = StepStatus.PENDING if enabled else StepStatus.DISABLED

    def get_step(self, step_id: str) -> Optional[WorkflowStep]:
        """Get a step by ID."""
        return self.steps.get(step_id)

    def get_steps(self) -> List[WorkflowStep]:
        """Get all steps in order."""
        return [self.steps[step_id] for step_id in self.step_order]

    def reset(self) -> None:
        """Reset all steps to initial state."""
        for step in self.steps.values():
            step.reset()
        self.current_step_id = None
        self.should_stop = False

    def _create_run_directory(self) -> Path:
        """Create a timestamped run directory."""
        timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
        run_dir = self.output_base_dir / f"{timestamp}"
        run_dir.mkdir(parents=True, exist_ok=True)
        return run_dir

    def _create_step_directory(self, step_index: int, step: WorkflowStep) -> Path:
        """Create a directory for a specific step."""
        step_dir = self.run_directory / f"{step_index:02d}_{step.id}"
        step_dir.mkdir(parents=True, exist_ok=True)
        return step_dir

    def log(self, message: str, level: str = "INFO") -> None:
        """Log a message."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted = f"[{timestamp}] [{level}] {message}"
        self.log_queue.put((formatted, level))
        if self.on_log:
            self.on_log(formatted, level)

    def stop(self) -> None:
        """Request the workflow to stop after the current step."""
        self.should_stop = True
        self.log("Workflow stop requested", "WARNING")

    def run(self, context: Optional[Dict[str, Any]] = None) -> Dict[str, StepResult]:
        """
        Execute all enabled steps in the workflow.

        Args:
            context: Optional dictionary of parameters passed to each step.

        Returns:
            Dictionary mapping step IDs to their results.
        """
        if self.is_running:
            raise RuntimeError("Workflow is already running")

        self.is_running = True
        self.should_stop = False
        results: Dict[str, StepResult] = {}
        context = context or {}

        # Create run directory
        self.run_directory = self._create_run_directory()
        context['run_directory'] = self.run_directory
        self.log(f"Starting workflow '{self.name}'", "INFO")
        self.log(f"Output directory: {self.run_directory}", "INFO")

        try:
            total_steps = len([s for s in self.get_steps() if s.enabled])
            completed_steps = 0

            for index, step_id in enumerate(self.step_order, 1):
                if self.should_stop:
                    self.log("Workflow stopped by user", "WARNING")
                    break

                step = self.steps[step_id]

                # Skip disabled steps
                if not step.enabled:
                    step.status = StepStatus.DISABLED
                    results[step_id] = StepResult(
                        success=True,
                        message="Step disabled"
                    )
                    continue

                # Check dependencies
                deps_ok = True
                for dep_id in step.dependencies:
                    dep_result = results.get(dep_id)
                    if not dep_result or not dep_result.success:
                        deps_ok = False
                        self.log(f"Skipping '{step.name}': dependency '{dep_id}' failed", "WARNING")
                        break

                if not deps_ok:
                    step.status = StepStatus.SKIPPED
                    results[step_id] = StepResult(
                        success=False,
                        message="Skipped due to failed dependency"
                    )
                    continue

                # Create step directory if enabled
                if self.create_step_directories:
                    step_dir = self._create_step_directory(index, step)
                    context['step_directory'] = step_dir

                # Execute step
                self.current_step_id = step_id
                step.status = StepStatus.IN_PROGRESS
                step.started_at = datetime.now()

                if self.on_step_start:
                    self.on_step_start(step)

                self.log(f"Starting step: {step.name}", "INFO")

                try:
                    result = step.function(context)
                    step.result = result

                    if result.success:
                        step.status = StepStatus.COMPLETED
                        self.log(f"Completed: {step.name} - {result.message}", "SUCCESS")
                    else:
                        step.status = StepStatus.ERROR
                        self.log(f"Failed: {step.name} - {result.message}", "ERROR")

                except Exception as e:
                    error_msg = f"{type(e).__name__}: {str(e)}"
                    step.result = StepResult(
                        success=False,
                        message=error_msg,
                        error=e
                    )
                    step.status = StepStatus.ERROR
                    self.log(f"Error in '{step.name}': {error_msg}", "ERROR")
                    self.log(traceback.format_exc(), "DEBUG")

                step.completed_at = datetime.now()
                if step.started_at:
                    step.result.duration = (step.completed_at - step.started_at).total_seconds()

                results[step_id] = step.result

                if self.on_step_complete:
                    self.on_step_complete(step)

                completed_steps += 1
                if self.on_progress:
                    self.on_progress(completed_steps, total_steps)

                # Check if should continue after error
                if step.status == StepStatus.ERROR:
                    if not (step.continue_on_error or self.continue_on_error):
                        self.log("Workflow stopped due to error", "ERROR")
                        break

        finally:
            self.is_running = False
            self.current_step_id = None

        # Summary
        success_count = sum(1 for r in results.values() if r.success)
        self.log(f"Workflow completed: {success_count}/{len(results)} steps successful", "INFO")

        return results

    def run_async(self, context: Optional[Dict[str, Any]] = None,
                  callback: Optional[Callable[[Dict[str, StepResult]], None]] = None) -> threading.Thread:
        """Run the workflow in a separate thread."""
        def _run():
            results = self.run(context)
            if callback:
                callback(results)

        thread = threading.Thread(target=_run, daemon=True)
        thread.start()
        return thread


def create_step(
    step_id: str,
    name: str,
    description: str,
    continue_on_error: bool = False
) -> Callable:
    """Decorator to create a workflow step from a function."""
    def decorator(func: Callable[..., StepResult]) -> WorkflowStep:
        return WorkflowStep(
            id=step_id,
            name=name,
            description=description,
            function=func,
            continue_on_error=continue_on_error
        )
    return decorator
