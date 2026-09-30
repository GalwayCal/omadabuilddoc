"""Pulls everything the as-built needs out of the controller into one plain dict.

The dict is also written to disk as JSON so a report can be re-rendered
(``--from-json``) without touching the controller again, and so you can see
exactly what the API returned when something looks wrong.

Every call is isolated: a failure (missing permission, endpoint not present on
this controller build, device offline) is recorded under ``site["errors"]``
and the rest of the collection carries on.
"""
from __future__ import annotations

import base64
import datetime as dt
import logging
from typing import Any, Callable

from .client import OpenApiClient, WebApiClient

log = logging.getLogger(__name__)

# Web-API paths tried for floor-plan maps (not in the Open API).  If none
# works on your build, open the Map page with browser DevTools > Network,
# find the XHR that lists maps and pass it with --maps-path.
DEFAULT_MAP_PATHS = ["sites/{siteId}/maps", "sites/{siteId}/map/maps", "sites/{siteId}/setting/maps"]


class Collector:
    def __init__(self, api: OpenApiClient, web: WebApiClient | None = None,
                 site_filter: list[str] | None = None, maps_path: str | None = None):
        self.api = api
        self.web = web
        self.site_filter = [s.lower() for s in (site_filter or [])]
        self.map_paths = [maps_path] if maps_path else DEFAULT_MAP_PATHS

    # ---------------------------------------------------------------- utils
    @staticmethod
    def _try(site: dict, what: str, fn: Callable[[], Any], default: Any = None) -> Any:
        try:
            return fn()
        except Exception as e:  # noqa: BLE001 - record and continue
            log.warning("  %-28s FAILED: %s", what, e)
            site.setdefault("errors", []).append({"what": what, "error": str(e)})
            return default

    def _first(self, site: dict, what: str, candidates, default=None):
        def run():
            res, ep = self.api.get_first(candidates)
            site.setdefault("endpoints", {})[what] = ep
            return res
        return self._try(site, what, run, default)

    # ----------------------------------------------------------- top level
    def collect(self) -> dict:
        api = self.api
        out: dict = {
            "generated": dt.datetime.now().astimezone().isoformat(timespec="seconds"),
            "controller": {"url": api.base_url, "omadacId": api.omadac_id, "info": api.info},
            "sites": [],
            "errors": [],
        }
        ctl = out["controller"]
        ctl["status"] = self._try(out, "controller status",
                                  lambda: api.get("system/setting/controller-status"), {})
        ctl["general"] = self._try(out, "controller general settings",
                                   lambda: api.get("global/controller/setting/general"), {})
        if not ctl["info"]:
            ctl["info"] = self._try(out, "controller /api/info", api.controller_info, {})

        sites = self._try(out, "site list", lambda: api.get_paged("sites"), []) or []
        log.info("Found %d site(s)", len(sites))
        for row in sites:
            if self.site_filter and row.get("name", "").lower() not in self.site_filter \
                    and row.get("siteId") not in self.site_filter:
                continue
            out["sites"].append(self.collect_site(row))
        return out

    # ---------------------------------------------------------------- site
    def collect_site(self, row: dict) -> dict:
        api, sid = self.api, row["siteId"]
        log.info("Collecting site '%s' (%s)", row.get("name"), sid)
        site: dict = {"siteId": sid, "name": row.get("name"), "errors": [], "endpoints": {}}
        base = f"sites/{sid}"

        detail = self._try(site, "site detail", lambda: api.get(base), {}) or {}
        site["info"] = {**row, **detail}

        # --- devices -----------------------------------------------------
        site["devices"] = self._try(site, "device list",
                                    lambda: api.get_paged(f"{base}/devices"), []) or []
        details: dict = {}
        for d in site["devices"]:
            mac, typ = d.get("mac"), (d.get("type") or "").lower()
            if not mac:
                continue
            if typ == "switch":
                details[mac] = {"switch": self._try(site, f"switch {d.get('name')}",
                                                    lambda: api.get(f"{base}/switches/{mac}"))}
            elif typ == "gateway":
                details[mac] = {
                    "gateway": self._try(site, f"gateway {d.get('name')}",
                                         lambda: api.get(f"{base}/gateways/{mac}")),
                    "wan_status": self._try(site, f"gateway WAN status {d.get('name')}",
                                            lambda: api.get(f"{base}/gateways/{mac}/wan-status")),
                }
            elif typ == "ap":
                details[mac] = {"ap": self._try(site, f"AP {d.get('name')}",
                                                lambda: api.get(f"{base}/aps/{mac}"))}
        site["device_detail"] = details

        # --- topology ----------------------------------------------------
        site["topology"] = self._first(site, "topology", [(f"{base}/topology", "v2", False)])

        # --- wired networks / VLANs --------------------------------------
        site["lan_networks"] = self._first(site, "LAN networks", [
            (f"{base}/lan-networks", "v2", True),
            (f"{base}/lan-networks", "v1", True),
        ], []) or []
        site["lan_profiles"] = self._try(site, "LAN (port) profiles",
                                         lambda: api.get_paged(f"{base}/lan-profiles"), []) or []
        site["internet"] = self._try(site, "internet / WAN config",
                                     lambda: api.get(f"{base}/internet"), {}) or {}
        site["dhcp_reservations"] = self._try(
            site, "DHCP reservations",
            lambda: api.get_paged(f"{base}/setting/service/dhcp"), []) or []

        # --- wireless ----------------------------------------------------
        site["radius_profiles"] = self._try(site, "RADIUS profiles",
                                            lambda: api.get_paged(f"{base}/profiles/radius"), []) or []
        site["ppsk_profiles"] = self._try(site, "PPSK profiles",
                                          lambda: api.get_paged(f"{base}/ppsk-profiles"), []) or []
        wlans = self._try(site, "WLAN groups",
                          lambda: api.get(f"{base}/wireless-network/wlans"), []) or []
        if isinstance(wlans, dict):
            wlans = wlans.get("data", [])
        for w in wlans:
            wid = w.get("wlanId")
            ssids = self._try(site, f"SSIDs in {w.get('name')}", lambda: api.get_paged(
                f"{base}/wireless-network/wlans/{wid}/ssids"), []) or []
            full = []
            for s in ssids:
                d = self._try(site, f"SSID detail {s.get('name')}", lambda: api.get(
                    f"{base}/wireless-network/wlans/{wid}/ssids/{s.get('ssidId')}"), None)
                full.append({**s, **(d or {})})
            w["ssids"] = full
        site["wlan_groups"] = wlans

        # --- maps (web API, optional) ------------------------------------
        site["maps"] = self._collect_maps(site) if self.web else []
        return site

    def _collect_maps(self, site: dict) -> list:
        sid = site["siteId"]
        for p in self.map_paths:
            path = p.format(siteId=sid)
            try:
                res = self.web.get(path, currentPage=1, currentPageSize=100)
            except Exception as e:  # noqa: BLE001
                log.debug("map path %s failed: %s", path, e)
                continue
            rows = res.get("data", res) if isinstance(res, dict) else res
            if not isinstance(rows, list):
                continue
            site["endpoints"]["maps (web API)"] = f"/{{omadacId}}/api/v2/{p}"
            for m in rows:
                img_ref = next((v for k, v in m.items() if isinstance(v, str) and
                                any(t in k.lower() for t in ("url", "image", "img", "picture"))), None)
                if img_ref:
                    try:
                        m["_image_b64"] = base64.b64encode(self.web.get_bytes(img_ref)).decode()
                    except Exception as e:  # noqa: BLE001
                        site["errors"].append({"what": f"map image {m.get('name')}", "error": str(e)})
            return rows
        site["errors"].append({"what": "floor-plan maps",
                               "error": "no web-API map endpoint answered; use --floorplans or --maps-path"})
        return []
