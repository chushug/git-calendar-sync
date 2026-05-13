from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime


@dataclass(frozen=True)
class Commit:
    hash: str
    email: str
    dt: datetime        # UTC-aware (committer date)
    subject: str
    branch: str = ""    # branch name at sync time
    remote_url: str = ""

    @property
    def short_hash(self) -> str:
        return self.hash[:7]

    @property
    def uid(self) -> str:
        """Stable RFC 5545 UID.  Never changes for the same commit hash."""
        slug = ""
        if self.remote_url:
            slug = self.remote_url.rstrip("/").split("/")[-1].removesuffix(".git")
        elif self.branch:
            slug = self.branch
        else:
            slug = "local"
        return f"{self.hash}@{slug}.commitcal"

    def __str__(self) -> str:
        ts = self.dt.strftime("%Y-%m-%d %H:%M")
        return f"[{self.short_hash}] {ts}  {self.subject}"
