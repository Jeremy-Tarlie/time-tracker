"""État et helpers du chronomètre (sans UI)."""

from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime

from time_tracker.constants import ActiveSession
from time_tracker.session_store import clear_active_session, save_active_session


def format_elapsed(seconds: float) -> str:
    total = max(0, int(seconds))
    hours, rem = divmod(total, 3600)
    minutes, secs = divmod(rem, 60)
    return f"{hours:02}:{minutes:02}:{secs:02}"


@dataclass
class TimerState:
    tracking: bool = False
    paused: bool = False
    active_project: str | None = None
    start_time: datetime | None = None
    elapsed: float = 0.0
    accumulated: float = 0.0
    running: bool = False
    timer_gen: int = 0
    writing: bool = False

    def bump_gen(self) -> int:
        self.timer_gen += 1
        return self.timer_gen

    def begin(self, project: str, *, now: datetime | None = None) -> int:
        moment = now or datetime.now()
        self.tracking = True
        self.paused = False
        self.active_project = project
        self.start_time = moment
        self.accumulated = 0.0
        self.elapsed = 0.0
        self.running = True
        return self.bump_gen()

    def resume_from_pause(self, *, now: datetime | None = None) -> int:
        moment = now or datetime.now()
        self.paused = False
        self.start_time = moment
        self.running = True
        return self.bump_gen()

    def pause(self) -> int:
        self.paused = True
        self.running = False
        self.accumulated = self.elapsed
        return self.bump_gen()

    def stop_clock(self) -> tuple[str, datetime, datetime, float]:
        """Arrête le chrono ; retourne project, session_start, end, elapsed."""
        self.running = False
        self.bump_gen()
        if not self.paused and self.start_time is not None:
            self.elapsed = self.accumulated + (datetime.now() - self.start_time).total_seconds()
        else:
            self.elapsed = self.accumulated
        end_time = datetime.now()
        project = self.active_project or ""
        session_start = datetime.fromtimestamp(end_time.timestamp() - self.elapsed)
        elapsed = self.elapsed
        clear_active_session()
        self.tracking = False
        self.paused = False
        self.active_project = None
        return project, session_start, end_time, elapsed

    def tick_elapsed(self) -> float:
        if self.start_time is None:
            return self.elapsed
        self.elapsed = self.accumulated + (datetime.now() - self.start_time).total_seconds()
        return self.elapsed

    def persist(self) -> None:
        if not self.tracking or self.active_project is None or self.start_time is None:
            clear_active_session()
            return
        if self.paused:
            accumulated = self.accumulated
            resume_at = None
        else:
            accumulated = self.elapsed
            resume_at = datetime.now().isoformat(timespec="seconds")
        session: ActiveSession = {
            "project": self.active_project,
            "start": self.start_time.isoformat(timespec="seconds"),
            "accumulated_seconds": accumulated,
            "paused": self.paused,
            "resume_at": resume_at,
        }
        save_active_session(session)

    def restore_running(
        self,
        project: str,
        elapsed: float,
        *,
        now: datetime | None = None,
    ) -> int:
        moment = now or datetime.now()
        self.active_project = project
        self.tracking = True
        self.paused = False
        self.accumulated = elapsed
        self.start_time = moment
        self.elapsed = elapsed
        self.running = True
        return self.bump_gen()

    def reset_idle(self) -> None:
        self.tracking = False
        self.paused = False
        self.active_project = None
        self.start_time = None
        self.accumulated = 0.0
        self.elapsed = 0.0
        self.running = False
