"""HTTP clients for a self-hosted Omada Software Controller (v5.12+ / v6.x).

Two clients live here:

* ``OpenApiClient`` - the official Omada Open API (OAuth2 client-credentials).
  This is what almost everything uses.  Create the app under
  Global View > Settings > Platform Integration > Open API (mode: Client),
  give it at least *Viewer* access to the sites you want documented.

* ``WebApiClient`` - the controller's internal web UI API (username/password).
  Only used for things the Open API does not expose (floor-plan Maps).
  It is optional and best-effort because TP-Link changes it between releases.
"""
from __future__ import annotations

import logging
import time
from typing import Any, Iterable

import requests
import urllib3

log = logging.getLogger(__name__)

TOKEN_EXPIRED_CODES = {-44112, -44113}  # expired / invalid access token


class OmadaError(RuntimeError):
    def __init__(self, path: str, code: Any, msg: str):
        super().__init__(f"{path}: errorCode={code} {msg}")
        self.path, self.code, self.msg = path, code, msg


class _Base:
    def __init__(self, base_url: str, verify_ssl: bool | str = True, timeout: int = 30):
        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self.s = requests.Session()
        self.s.verify = verify_ssl
        if verify_ssl is False:
            urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

    def controller_info(self) -> dict:
        """GET /api/info - unauthenticated; returns omadacId and controllerVer."""
        r = self.s.get(f"{self.base_url}/api/info", timeout=self.timeout)
        r.raise_for_status()
        data = r.json()
        return data.get("result", data)


class OpenApiClient(_Base):
    def __init__(self, base_url: str, client_id: str, client_secret: str,
                 omadac_id: str | None = None, verify_ssl: bool | str = True,
                 timeout: int = 30, page_size: int = 1000):
        super().__init__(base_url, verify_ssl, timeout)
        self.client_id = client_id
        self.client_secret = client_secret
        self.page_size = page_size
        self.info: dict = {}
        if not omadac_id:
            self.info = self.controller_info()
            omadac_id = self.info.get("omadacId")
            if not omadac_id:
                raise RuntimeError("Could not discover omadacId from /api/info; set it explicitly")
        self.omadac_id = omadac_id
        self._token: str | None = None
        self._token_expiry = 0.0

    # ---------------------------------------------------------------- auth
    def _authenticate(self) -> None:
        url = f"{self.base_url}/openapi/authorize/token"
        body = {"omadacId": self.omadac_id, "client_id": self.client_id,
                "client_secret": self.client_secret}
        r = self.s.post(url, params={"grant_type": "client_credentials"}, json=body,
                        timeout=self.timeout)
        r.raise_for_status()
        data = r.json()
        if data.get("errorCode") != 0:
            raise OmadaError("/openapi/authorize/token", data.get("errorCode"), data.get("msg", ""))
        res = data["result"]
        self._token = res["accessToken"]
        self._token_expiry = time.time() + int(res.get("expiresIn", 3600)) - 60
        log.debug("Obtained Open API access token (expires in %ss)", res.get("expiresIn"))

    def _headers(self) -> dict:
        if not self._token or time.time() >= self._token_expiry:
            self._authenticate()
        return {"Authorization": f"AccessToken={self._token}",
                "Content-Type": "application/json"}

    # ------------------------------------------------------------ requests
    def url(self, path: str, version: str = "v1") -> str:
        """``path`` is relative to /openapi/{version}/{omadacId}."""
        return f"{self.base_url}/openapi/{version}/{self.omadac_id}/{path.lstrip('/')}"

    def request(self, method: str, path: str, version: str = "v1",
                params: dict | None = None, json: Any = None) -> Any:
        full = self.url(path, version)
        for attempt in (1, 2):
            r = self.s.request(method, full, headers=self._headers(), params=params,
                               json=json, timeout=self.timeout)
            if r.status_code == 401 and attempt == 1:
                self._token = None
                continue
            r.raise_for_status()
            data = r.json()
            code = data.get("errorCode", 0)
            if code in TOKEN_EXPIRED_CODES and attempt == 1:
                self._token = None
                continue
            if code != 0:
                raise OmadaError(f"/openapi/{version}/.../{path}", code, data.get("msg", ""))
            return data.get("result")
        raise OmadaError(path, "auth", "token refresh failed")

    def get(self, path: str, version: str = "v1", **params) -> Any:
        return self.request("GET", path, version, params=params or None)

    def get_paged(self, path: str, version: str = "v1", **params) -> list:
        """Walk page/pageSize pagination and return the concatenated ``data`` list."""
        out: list = []
        page = 1
        while True:
            res = self.get(path, version, page=page, pageSize=self.page_size, **params)
            if isinstance(res, list):  # some endpoints ignore paging
                return res
            rows = (res or {}).get("data") or []
            out.extend(rows)
            total = (res or {}).get("totalRows", len(out))
            if not rows or len(out) >= total:
                return out
            page += 1

    def get_first(self, candidates: Iterable[tuple[str, str, bool]]) -> tuple[Any, str]:
        """Try (path, version, paged) candidates in order; return first that works.

        Lets the tool cope with endpoints that moved between v1/v2 across
        controller releases.  Raises the last error if none succeed.
        """
        last: Exception | None = None
        for path, version, paged in candidates:
            try:
                res = self.get_paged(path, version) if paged else self.get(path, version)
                return res, f"/openapi/{version}/{{omadacId}}/{path}"
            except (OmadaError, requests.HTTPError) as e:
                log.debug("candidate %s %s failed: %s", version, path, e)
                last = e
        raise last or RuntimeError("no candidates")

    def api_spec(self) -> dict:
        """The controller publishes its full OpenAPI spec unauthenticated."""
        for p in ("/v3/api-docs", "/openapi/v3/api-docs"):
            try:
                r = self.s.get(self.base_url + p, timeout=self.timeout)
                if r.ok and r.headers.get("content-type", "").startswith("application/json"):
                    spec = r.json()
                    if isinstance(spec, dict) and spec.get("paths"):
                        return spec
            except requests.RequestException:
                pass
        return {}


class WebApiClient(_Base):
    """Minimal session for the internal web API (used for floor-plan Maps only)."""

    def __init__(self, base_url: str, username: str, password: str, omadac_id: str,
                 verify_ssl: bool | str = True, timeout: int = 30):
        super().__init__(base_url, verify_ssl, timeout)
        self.omadac_id = omadac_id
        r = self.s.post(f"{self.base_url}/{omadac_id}/api/v2/login",
                        json={"username": username, "password": password}, timeout=timeout)
        r.raise_for_status()
        data = r.json()
        if data.get("errorCode") != 0:
            raise OmadaError("/api/v2/login", data.get("errorCode"), data.get("msg", ""))
        self.token = data["result"]["token"]
        self.s.headers["Csrf-Token"] = self.token

    def get(self, path: str, **params) -> Any:
        params = {"token": self.token, **params}
        r = self.s.get(f"{self.base_url}/{self.omadac_id}/api/v2/{path.lstrip('/')}",
                       params=params, timeout=self.timeout)
        r.raise_for_status()
        data = r.json()
        if data.get("errorCode") != 0:
            raise OmadaError(path, data.get("errorCode"), data.get("msg", ""))
        return data.get("result")

    def get_bytes(self, path_or_url: str) -> bytes:
        url = path_or_url if path_or_url.startswith("http") else \
            f"{self.base_url}/{path_or_url.lstrip('/')}"
        r = self.s.get(url, params={"token": self.token}, timeout=self.timeout)
        r.raise_for_status()
        return r.content
