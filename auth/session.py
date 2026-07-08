from __future__ import annotations

from dataclasses import dataclass, field
from datetime import datetime, timezone
from typing import Any


@dataclass
class AuthSession:
	access_token: str | None = None
	refresh_token: str | None = None
	account: dict[str, Any] = field(default_factory=dict)
	expires_at: datetime | None = None

	def is_authenticated(self) -> bool:
		return bool(self.access_token) and (self.expires_at is None or self.expires_at > datetime.now(timezone.utc))


ACTIVE_SESSION = AuthSession()


def update_session(
	*,
	access_token: str | None = None,
	refresh_token: str | None = None,
	account: dict[str, Any] | None = None,
	expires_at: datetime | None = None,
) -> AuthSession:
	if access_token is not None:
		ACTIVE_SESSION.access_token = access_token
	if refresh_token is not None:
		ACTIVE_SESSION.refresh_token = refresh_token
	if account is not None:
		ACTIVE_SESSION.account = account
	if expires_at is not None:
		ACTIVE_SESSION.expires_at = expires_at
	return ACTIVE_SESSION


def clear_session() -> None:
	ACTIVE_SESSION.access_token = None
	ACTIVE_SESSION.refresh_token = None
	ACTIVE_SESSION.account = {}
	ACTIVE_SESSION.expires_at = None
