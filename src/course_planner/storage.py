"""Microsoft Graph-backed SharePoint file storage.

The adapter is disabled by default. Simple Graph uploads are limited to 250 MiB.
"""

from __future__ import annotations

import json
import math
import os
import time
import urllib.error
import urllib.parse
import urllib.request
from collections.abc import Mapping
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime


class SharePointStorageError(RuntimeError):
    """Base exception for SharePoint storage errors."""


class StorageDisabledError(SharePointStorageError):
    """Raised when storage operations are attempted while the adapter is disabled."""


class StorageConfigurationError(SharePointStorageError):
    """Raised when enabled storage is missing valid configuration."""


class _HTTPStatusError(Exception):
    def __init__(self, status: int, retry_after: str | None = None) -> None:
        self.status = status
        self.retry_after = retry_after
        super().__init__(str(status))


class _SafeRedirectHandler(urllib.request.HTTPRedirectHandler):
    """Keep redirects HTTPS-only and prevent credentials crossing authority boundaries."""

    def __init__(self, *, block_redirects: bool = False) -> None:
        super().__init__()
        self.block_redirects = block_redirects

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        if self.block_redirects:
            return None
        try:
            destination = urllib.parse.urlsplit(newurl)
            source = urllib.parse.urlsplit(req.full_url)
            if destination.scheme.lower() != "https":
                raise SharePointStorageError("SharePoint redirect was rejected.")
            if destination.username is not None or destination.password is not None:
                raise SharePointStorageError("SharePoint redirect was rejected.")
        except ValueError:
            raise SharePointStorageError("SharePoint redirect was rejected.") from None

        redirected = super().redirect_request(req, fp, code, msg, headers, newurl)
        if redirected is None:
            return None
        if (source.hostname, source.port) != (destination.hostname, destination.port):
            redirected.remove_header("Authorization")
            redirected.unredirected_hdrs.pop("Authorization", None)
        return redirected


class GraphStorage:
    """Read and write files in configured SharePoint drive folders via Graph."""

    MAX_UPLOAD_BYTES = 250 * 1024 * 1024
    DEFAULT_TIMEOUT = 20.0
    MAX_TIMEOUT = 60.0
    DEFAULT_RETRIES = 3
    MAX_RETRIES = 3
    MAX_RETRY_DELAY = 5.0

    def __init__(
        self,
        tenant_id: str = "",
        client_id: str = "",
        client_secret: str = "",
        drive_id: str = "",
        input_root: str = "",
        output_root: str = "",
        *,
        enabled: bool = True,
        timeout: float = DEFAULT_TIMEOUT,
        retries: int = DEFAULT_RETRIES,
    ) -> None:
        self.enabled = enabled
        self.timeout = min(float(timeout), self.MAX_TIMEOUT)
        self.retries = retries
        self._access_token: str | None = None
        self._token_expires_at = 0.0

        if (
            not math.isfinite(self.timeout)
            or self.timeout <= 0
            or not isinstance(retries, int)
            or isinstance(retries, bool)
            or not 0 <= retries <= self.MAX_RETRIES
        ):
            raise StorageConfigurationError("Invalid storage timeout or retry configuration.")

        if not enabled:
            self.tenant_id = ""
            self.client_id = ""
            self.client_secret = ""
            self.drive_id = ""
            self.input_root = ""
            self.output_root = ""
            return

        required = {
            "GRAPH_TENANT_ID": tenant_id,
            "GRAPH_CLIENT_ID": client_id,
            "GRAPH_CLIENT_SECRET": client_secret,
            "GRAPH_DRIVE_ID": drive_id,
        }
        missing = [name for name, value in required.items() if not value.strip()]
        if missing:
            raise StorageConfigurationError(
                "Missing required SharePoint setting(s): " + ", ".join(missing) + "."
            )

        self.tenant_id = tenant_id
        self.client_id = client_id
        self.client_secret = client_secret
        self.drive_id = drive_id
        self.input_root = self._validate_root(input_root, "GRAPH_INPUT_ROOT")
        self.output_root = self._validate_root(output_root, "GRAPH_OUTPUT_ROOT")

    @classmethod
    def from_env(cls, environ: Mapping[str, str] | None = None) -> GraphStorage:
        """Create an adapter from environment settings; disabled unless explicitly enabled."""
        env = os.environ if environ is None else environ
        enabled_value = env.get("PLANNER_SHAREPOINT_ENABLED", "false").strip().lower()
        if enabled_value not in {"true", "false", "1", "0", "yes", "no", "on", "off"}:
            raise StorageConfigurationError("PLANNER_SHAREPOINT_ENABLED must be a boolean value.")
        enabled = enabled_value in {"true", "1", "yes", "on"}
        if not enabled:
            return cls(enabled=False)

        names = (
            "GRAPH_TENANT_ID",
            "GRAPH_CLIENT_ID",
            "GRAPH_CLIENT_SECRET",
            "GRAPH_DRIVE_ID",
            "GRAPH_INPUT_ROOT",
            "GRAPH_OUTPUT_ROOT",
        )
        missing = [name for name in names if name not in env]
        if missing:
            raise StorageConfigurationError(
                "Missing required SharePoint setting(s): " + ", ".join(missing) + "."
            )
        return cls(
            tenant_id=env["GRAPH_TENANT_ID"],
            client_id=env["GRAPH_CLIENT_ID"],
            client_secret=env["GRAPH_CLIENT_SECRET"],
            drive_id=env["GRAPH_DRIVE_ID"],
            input_root=env["GRAPH_INPUT_ROOT"],
            output_root=env["GRAPH_OUTPUT_ROOT"],
        )

    def download(self, relative_path: str) -> bytes:
        """Download a file below GRAPH_INPUT_ROOT and return its bytes."""
        self._ensure_enabled()
        path = self._combined_path(self.input_root, relative_path)
        return self._graph_request("GET", self._content_url(path))

    def upload(self, relative_path: str, content: bytes) -> None:
        """Upload bytes below GRAPH_OUTPUT_ROOT using Graph's simple upload endpoint."""
        self._ensure_enabled()
        if not isinstance(content, bytes):
            raise TypeError("content must be bytes.")
        if len(content) > self.MAX_UPLOAD_BYTES:
            raise SharePointStorageError(
                f"Upload exceeds the {self.MAX_UPLOAD_BYTES}-byte simple-upload limit."
            )
        path = self._combined_path(self.output_root, relative_path)
        relative_directory = relative_path.rpartition("/")[0]
        self.ensure_folder(relative_directory)
        self._graph_request("PUT", self._content_url(path), content)

    def ensure_folder(self, relative_path: str) -> None:
        """Create each missing folder below GRAPH_OUTPUT_ROOT.

        A 409 is accepted only when Graph confirms that the conflicting item is a folder.
        An empty relative path ensures GRAPH_OUTPUT_ROOT itself.
        """
        self._ensure_enabled()
        try:
            relative_parts = self._validate_relative_path(relative_path) if relative_path else ()
        except ValueError:
            raise SharePointStorageError("Folder path must be a safe relative path.") from None
        root_parts = self._validate_relative_path(self.output_root) if self.output_root else ()
        parts = (*root_parts, *relative_parts)
        if not parts:
            return

        parent_id: str | None = None
        created_path: list[str] = []
        for part in parts:
            created_path.append(part)
            payload = json.dumps(
                {
                    "name": part,
                    "folder": {},
                    "@microsoft.graph.conflictBehavior": "fail",
                }
            ).encode("utf-8")
            body = self._graph_request(
                "POST",
                self._children_url(parent_id),
                payload,
                content_type="application/json",
                allowed_status=(409,),
            )
            item = (
                self._parse_folder_item(body)
                if body
                else self._get_existing_folder("/".join(created_path))
            )
            parent_id = item["id"]

    def _ensure_enabled(self) -> None:
        if not self.enabled:
            raise StorageDisabledError("SharePoint storage is disabled.")

    @staticmethod
    def _validate_root(root: str, setting_name: str) -> str:
        if not isinstance(root, str):
            raise StorageConfigurationError(f"{setting_name} must be a relative folder path.")
        if not root:
            return ""
        try:
            GraphStorage._validate_relative_path(root)
        except ValueError:
            raise StorageConfigurationError(
                f"{setting_name} must be a safe relative folder path."
            ) from None
        return root

    @staticmethod
    def _validate_relative_path(path: str) -> tuple[str, ...]:
        if not isinstance(path, str) or not path or path.startswith(("/", "\\")):
            raise ValueError("invalid path")
        if "\\" in path or any(ord(char) < 32 or ord(char) == 127 for char in path):
            raise ValueError("invalid path")
        parts = tuple(path.split("/"))
        if any(part in {"", ".", ".."} for part in parts):
            raise ValueError("invalid path")
        return parts

    @classmethod
    def _combined_path(cls, root: str, relative_path: str) -> str:
        try:
            relative_parts = cls._validate_relative_path(relative_path)
        except ValueError:
            raise SharePointStorageError(
                "File path must be a safe, non-empty relative path."
            ) from None
        root_parts = cls._validate_relative_path(root) if root else ()
        return "/".join((*root_parts, *relative_parts))

    def _content_url(self, path: str) -> str:
        encoded_path = "/".join(urllib.parse.quote(part, safe="-._~") for part in path.split("/"))
        drive_id = urllib.parse.quote(self.drive_id, safe="-._~")
        return f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_path}:/content"

    def _children_url(self, parent_id: str | None) -> str:
        drive_id = urllib.parse.quote(self.drive_id, safe="-._~")
        if parent_id is None:
            return f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root/children"
        encoded_id = urllib.parse.quote(parent_id, safe="-._~")
        return f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{encoded_id}/children"

    def _item_url(self, path: str) -> str:
        encoded_path = "/".join(urllib.parse.quote(part, safe="-._~") for part in path.split("/"))
        drive_id = urllib.parse.quote(self.drive_id, safe="-._~")
        return f"https://graph.microsoft.com/v1.0/drives/{drive_id}/root:/{encoded_path}"

    def _get_existing_folder(self, path: str) -> dict[str, str]:
        try:
            body = self._graph_request("GET", self._item_url(path))
        except SharePointStorageError:
            raise SharePointStorageError(
                "SharePoint folder conflict could not be verified."
            ) from None
        return self._parse_folder_item(body)

    @staticmethod
    def _parse_folder_item(body: bytes) -> dict[str, str]:
        try:
            item = json.loads(body)
            item_id = item["id"]
            folder = item["folder"]
            if not isinstance(item_id, str) or not item_id or not isinstance(folder, dict):
                raise ValueError("invalid folder metadata")
        except (KeyError, TypeError, ValueError):
            raise SharePointStorageError("SharePoint item is not a verifiable folder.") from None
        return {"id": item_id}

    def _token_url(self) -> str:
        tenant_id = urllib.parse.quote(self.tenant_id, safe="-._~")
        return f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"

    def _get_access_token(self, *, force_refresh: bool = False) -> str:
        if (
            not force_refresh
            and self._access_token is not None
            and time.monotonic() < self._token_expires_at
        ):
            return self._access_token

        form = urllib.parse.urlencode(
            {
                "client_id": self.client_id,
                "client_secret": self.client_secret,
                "scope": "https://graph.microsoft.com/.default",
                "grant_type": "client_credentials",
            }
        ).encode("ascii")
        request = urllib.request.Request(
            self._token_url(),
            data=form,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        try:
            response_body = self._send_with_retries(request)
            response = json.loads(response_body)
            token = response["access_token"]
            if not isinstance(token, str) or not token:
                raise ValueError("invalid token")
            expires_in = float(response.get("expires_in", 3600))
        except _HTTPStatusError as error:
            raise SharePointStorageError(
                f"SharePoint authentication failed (HTTP {error.status})."
            ) from None
        except (KeyError, TypeError, ValueError):
            raise SharePointStorageError(
                "SharePoint authentication response was invalid."
            ) from None

        self._access_token = token
        self._token_expires_at = time.monotonic() + max(0.0, expires_in - 60.0)
        return token

    def _graph_request(
        self,
        method: str,
        url: str,
        data: bytes | None = None,
        *,
        content_type: str = "application/octet-stream",
        allowed_status: tuple[int, ...] = (),
    ) -> bytes:
        for token_attempt in range(2):
            token = self._get_access_token(force_refresh=token_attempt == 1)
            headers = {"Authorization": f"Bearer {token}"}
            if data is not None:
                headers["Content-Type"] = content_type
            request = urllib.request.Request(url, data=data, headers=headers, method=method)
            try:
                return self._send_with_retries(request)
            except _HTTPStatusError as error:
                if error.status == 401 and token_attempt == 0:
                    self._access_token = None
                    continue
                if error.status in allowed_status:
                    return b""
                raise SharePointStorageError(
                    f"SharePoint file request failed (HTTP {error.status})."
                ) from None
        raise SharePointStorageError("SharePoint authentication failed.")

    def _send_with_retries(self, request: urllib.request.Request) -> bytes:
        token_post = (
            request.get_method() == "POST"
            and urllib.parse.urlsplit(request.full_url).hostname == "login.microsoftonline.com"
        )
        opener = urllib.request.build_opener(_SafeRedirectHandler(block_redirects=token_post))
        for attempt in range(self.retries + 1):
            try:
                with opener.open(request, timeout=self.timeout) as response:
                    return response.read()
            except urllib.error.HTTPError as error:
                status = error.code
                retry_after = error.headers.get("Retry-After") if error.headers else None
                error.close()
                if status in {429, 503} and attempt < self.retries:
                    time.sleep(self._retry_delay(retry_after))
                    continue
                raise _HTTPStatusError(status, retry_after) from None
            except (urllib.error.URLError, TimeoutError, OSError):
                raise SharePointStorageError("SharePoint request failed.") from None
        raise SharePointStorageError("SharePoint request failed.")

    @classmethod
    def _retry_delay(cls, retry_after: str | None) -> float:
        if not retry_after:
            return 0.0
        try:
            delay = float(retry_after)
        except ValueError:
            try:
                retry_at = parsedate_to_datetime(retry_after)
                if retry_at.tzinfo is None:
                    retry_at = retry_at.replace(tzinfo=timezone.utc)
                delay = (retry_at - datetime.now(timezone.utc)).total_seconds()
            except (TypeError, ValueError, OverflowError):
                return 0.0
        return min(max(delay, 0.0), cls.MAX_RETRY_DELAY)
