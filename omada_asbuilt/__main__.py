"""Command line entry point:  python -m omada_asbuilt --help"""
from __future__ import annotations

import argparse
import configparser
import getpass
import json
import logging
import os
import re
import sys
from pathlib import Path

from .client import OpenApiClient, WebApiClient
from .collector import Collector
from .report import Report

log = logging.getLogger("omada_asbuilt")


def _cfg(args) -> dict:
    """Merge settings: CLI > environment (OMADA_*) > config file [omada] section."""
    cp = configparser.ConfigParser()
    if args.config and Path(args.config).exists():
        cp.read(args.config)
    sect = cp["omada"] if cp.has_section("omada") else {}

    def pick(name, env):
        v = getattr(args, name, None)
        return v if v not in (None, "") else os.environ.get(env) or sect.get(name)

    keys = {"url": "OMADA_URL", "client_id": "OMADA_CLIENT_ID", "client_secret": "OMADA_CLIENT_SECRET",
            "omadac_id": "OMADA_ID", "web_user": "OMADA_WEB_USER", "web_pass": "OMADA_WEB_PASS",
            "customer": "ASBUILT_CUSTOMER", "author": "ASBUILT_AUTHOR", "ca_bundle": "OMADA_CA_BUNDLE",
            "title": "ASBUILT_TITLE", "logo": "ASBUILT_LOGO"}
    c = {k: pick(k, e) for k, e in keys.items()}
    insecure = args.insecure or str(sect.get("insecure", "")).lower() in ("1", "true", "yes")
    c["verify"] = False if insecure else (c["ca_bundle"] or True)
    return c


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(prog="omada_asbuilt",
                                 description="Generate a Word as-built document from a TP-Link Omada controller.")
    ap.add_argument("-c", "--config", default="omada.ini", help="INI file with an [omada] section (default omada.ini)")
    ap.add_argument("--url", help="Controller base URL, e.g. https://omada.example.local:8043")
    ap.add_argument("--client-id", dest="client_id")
    ap.add_argument("--client-secret", dest="client_secret")
    ap.add_argument("--omadac-id", dest="omadac_id", help="Omada ID (auto-detected from /api/info if omitted)")
    ap.add_argument("--insecure", action="store_true", help="Skip TLS verification (self-signed controller cert)")
    ap.add_argument("--ca-bundle", dest="ca_bundle", help="PEM file to verify the controller certificate")
    ap.add_argument("--site", action="append", help="Only document this site name/ID (repeatable)")
    ap.add_argument("-o", "--out", help="Output .docx (default: <outdir>/Omada-AsBuilt-<date>.docx)")
    ap.add_argument("--outdir", default="output", help="Folder for the report and raw JSON (default ./output)")
    ap.add_argument("--title", help="Document title (default: controller name)")
    ap.add_argument("--customer", help="Customer / client name for the cover page")
    ap.add_argument("--author", help="Prepared-by name")
    ap.add_argument("--logo", help="PNG/JPG logo for the cover page")
    ap.add_argument("--secrets", choices=["show", "mask", "omit"], default="show",
                    help="How to print Wi-Fi/WAN passwords (default: show)")
    ap.add_argument("--floorplans", help="Folder of floor-plan images; files named <site name>*.png|jpg "
                                         "or placed in a sub-folder named after the site")
    ap.add_argument("--web-user", dest="web_user", help="Controller admin user for floor-plan maps (web API)")
    ap.add_argument("--web-pass", dest="web_pass", help="Password for --web-user (prompted if omitted)")
    ap.add_argument("--maps-path", help="Web API path listing maps, e.g. 'sites/{siteId}/maps'")
    ap.add_argument("--from-json", help="Re-render from a previously saved raw JSON (no controller access)")
    ap.add_argument("--probe", action="store_true", help="Test connectivity/auth, list sites, save API spec, exit")
    ap.add_argument("-v", "--verbose", action="store_true")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.DEBUG if args.verbose else logging.INFO,
                        format="%(levelname)-7s %(message)s")
    for noisy in ("urllib3", "matplotlib", "PIL"):
        logging.getLogger(noisy).setLevel(logging.WARNING)
    c = _cfg(args)
    outdir = Path(args.outdir)
    outdir.mkdir(parents=True, exist_ok=True)

    if args.from_json:
        data = json.loads(Path(args.from_json).read_text(encoding="utf-8"))
    else:
        missing = [k for k in ("url", "client_id", "client_secret") if not c[k]]
        if missing:
            ap.error(f"missing {', '.join(missing)} (use flags, OMADA_* env vars or {args.config})")
        api = OpenApiClient(c["url"], c["client_id"], c["client_secret"], c["omadac_id"], verify_ssl=c["verify"])
        log.info("Controller %s  omadacId=%s  version=%s", api.base_url, api.omadac_id,
                 api.info.get("controllerVer", "?"))
        if args.probe:
            return _probe(api, outdir)
        web = None
        if c["web_user"]:
            pw = c["web_pass"] or getpass.getpass(f"Password for {c['web_user']}: ")
            try:
                web = WebApiClient(c["url"], c["web_user"], pw, api.omadac_id, verify_ssl=c["verify"])
            except Exception as e:  # noqa: BLE001
                log.warning("Web API login failed (%s) - floor-plan maps will be skipped", e)
        data = Collector(api, web, args.site, args.maps_path).collect()
        raw = outdir / f"omada-raw-{data['generated'][:10]}.json"
        raw.write_text(json.dumps(data, indent=2, ensure_ascii=False), encoding="utf-8")
        log.info("Raw API data saved to %s", raw)

    st = (data.get("controller") or {}).get("status") or {}
    title = c["title"] or st.get("name") or "Omada Network"
    out = args.out or str(outdir / f"Omada-AsBuilt-{re.sub(r'[^A-Za-z0-9]+', '-', title).strip('-')}-"
                                   f"{data.get('generated', '')[:10]}.docx")
    Report(data, title=title, customer=c["customer"] or "", author=c["author"] or "",
           secrets=args.secrets, floorplans=args.floorplans, logo=c["logo"]).build(out)
    n_err = len(data.get("errors", [])) + sum(len(s.get("errors", [])) for s in data.get("sites", []))
    log.info("As-built written to %s  (%d site(s), %d item(s) not returned - see Appendix A)",
             out, len(data.get("sites", [])), n_err)
    return 0


def _probe(api: OpenApiClient, outdir: Path) -> int:
    sites = api.get_paged("sites")
    print(f"Auth OK. {len(sites)} site(s):")
    for s in sites:
        print(f"  - {s.get('name')}  ({s.get('siteId')})")
    spec = api.api_spec()
    if spec:
        p = outdir / "omada-openapi-spec.json"
        p.write_text(json.dumps(spec, indent=1), encoding="utf-8")
        print(f"Controller's OpenAPI spec ({len(spec.get('paths', {}))} paths) saved to {p}")
    else:
        print("Controller did not serve /v3/api-docs (open the API docs link in "
              "Settings > Platform Integration > Open API instead).")
    return 0


if __name__ == "__main__":
    sys.exit(main())
