from __future__ import annotations

from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any

from config.settings import MSAL_AUTHORITY, MSAL_CLIENT_ID, MSAL_TENANT_ID
from auth.session import update_session

try:
    import msal
except Exception:  # pragma: no cover - optional dependency during development
    msal = None


@dataclass
class LoginResult:
    success: bool
    account: dict[str, Any]
    access_token: str | None = None
    message: str = ""


class MicrosoftAuthenticator:
    def __init__(
        self,
        client_id: str | None = None,
        tenant_id: str | None = None,
        authority: str | None = None,
    ):
        self.client_id = client_id or MSAL_CLIENT_ID
        self.tenant_id = tenant_id or MSAL_TENANT_ID
        self.authority = authority or MSAL_AUTHORITY
        self.scopes = ["User.Read", "Files.Read.All", "Sites.Read.All"]
        self._app = None
        if msal and self.client_id:
            self._app = msal.PublicClientApplication(
                client_id=self.client_id, authority=self.authority
            )

    def is_configured(self) -> bool:
        return bool(self.client_id)

    def login_interactive(self, progress_callback=None) -> LoginResult:
        if self._app is None:
            if progress_callback:
                progress_callback(
                    {"status": "MSAL no está configurado, usando modo desarrollo."}
                )
            return self.login_mock()

        if progress_callback:
            progress_callback({"status": "Buscando sesión anterior..."})
        accounts = self._app.get_accounts()
        if accounts:
            if progress_callback:
                progress_callback({"status": "Intentando autenticación silenciosa..."})
            result = self._app.acquire_token_silent(self.scopes, account=accounts[0])
            if result and "access_token" in result:
                account = dict(accounts[0])
                token = result["access_token"]
                update_session(
                    access_token=token,
                    account=account,
                    expires_at=datetime.now(timezone.utc) + timedelta(hours=1),
                )
                return LoginResult(
                    True,
                    account=account,
                    access_token=token,
                    message="Authenticated silently",
                )

        if progress_callback:
            progress_callback({"status": "Iniciando Device Code Flow..."})
        flow = self._app.initiate_device_flow(scopes=self.scopes)
        if "user_code" not in flow:
            return LoginResult(
                False,
                account={},
                message=flow.get(
                    "error_description", "No se pudo iniciar el flujo de login."
                ),
            )

        if progress_callback:
            progress_callback(flow)

        result = self._app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            return LoginResult(
                False,
                account={},
                message=result.get("error_description", "Login cancelado o fallido."),
            )

        account = result.get("id_token_claims", {})
        token = result["access_token"]
        update_session(
            access_token=token,
            account=account,
            expires_at=datetime.now(timezone.utc) + timedelta(hours=1),
        )
        return LoginResult(
            True, account=account, access_token=token, message="Authenticated"
        )

    def login_mock(self) -> LoginResult:
        account = {
            "name": "Valeria Querebalu",
            "username": "vquerebalu@beta.com",
            "tenantId": self.tenant_id,
        }
        token = "mock-access-token"
        update_session(
            access_token=token,
            account=account,
            expires_at=datetime.now(timezone.utc) + timedelta(hours=8),
        )
        return LoginResult(
            True, account=account, access_token=token, message="Mock login active"
        )
