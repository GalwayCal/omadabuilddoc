"""Topology and logical-network diagrams rendered to PNG with matplotlib.

No Graphviz install needed: a small tree layout is done here.  Small sites are
drawn top-down; sites with many edge devices switch to left-to-right so the
image stays readable on a portrait A4/Letter page.
"""
from __future__ import annotations

import io
from collections import defaultdict

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402
from matplotlib.patches import FancyBboxPatch  # noqa: E402

COLORS = {  # fill, edge
    "internet": ("#EEF2F7", "#7A8699"),
    "gateway": ("#DCEBFA", "#1F5FA8"),
    "switch": ("#DFF2E6", "#1E7A45"),
    "ap": ("#FDF0DC", "#B06A00"),
    "olt": ("#EFE3F7", "#6B3FA0"),
    "network": ("#E8F1FB", "#2C6CB0"),
    "ssid": ("#FFF4E0", "#B06A00"),
    "other": ("#F1F1F1", "#666666"),
}
TYPE_LABEL = {"gateway": "Gateway", "switch": "Switch", "ap": "Access Point", "olt": "OLT"}
BOX_W, BOX_H = 2.5, 1.0


def _norm(mac: str | None) -> str:
    return (mac or "").upper().replace(":", "-")


# ------------------------------------------------------------------ graph
def build_graph(site: dict) -> tuple[dict, dict, dict]:
    """Return (nodes, parent_of, edge_meta).

    nodes: mac -> {name, type, model, ip, online}
    parent_of: child_mac -> parent_mac
    edge_meta: child_mac -> {"port": str, "wireless": bool, "speed": str}
    """
    nodes: dict = {}
    for d in site.get("devices", []):
        mac = _norm(d.get("mac"))
        nodes[mac] = {"name": d.get("name") or mac, "type": (d.get("type") or "other").lower(),
                      "model": d.get("model") or d.get("modelName") or "", "ip": d.get("ip") or "",
                      "online": d.get("status") in (1, 14, 15) or d.get("detailStatus") in (14, 15)}
    parent_of: dict = {}
    meta: dict = defaultdict(dict)

    topo = site.get("topology") or {}
    for n in topo.get("topologyNodes") or []:
        mac = _norm(n.get("mac"))
        nodes.setdefault(mac, {"name": n.get("name") or mac, "type": (n.get("type") or "other").lower(),
                               "model": n.get("model", ""), "ip": "", "online": True})
        if n.get("wirelessUpLink"):
            meta[mac]["wireless"] = True
    for e in topo.get("topologyEdges") or []:
        up, down = _norm(e.get("upLinkMac")), _norm(e.get("downLinkMac"))
        if up and down and up != down:
            parent_of[down] = up

    # Device list uplink info fills gaps and gives the port on the upstream device
    for d in site.get("devices", []):
        mac, up = _norm(d.get("mac")), _norm(d.get("uplinkDeviceMac"))
        if up and up != mac:
            parent_of.setdefault(mac, up)
            if d.get("uplinkDevicePort"):
                meta[mac]["port"] = str(d["uplinkDevicePort"])
        ap = (site.get("device_detail", {}).get(d.get("mac"), {}) or {}).get("ap") or {}
        if ap.get("wireless uplink info", {}).get("uplinkMac"):
            meta[mac]["wireless"] = True
    for child, par in list(parent_of.items()):  # drop links to unknown devices
        if par not in nodes:
            nodes[par] = {"name": par, "type": "other", "model": "", "ip": "", "online": True}
    return nodes, parent_of, dict(meta)


def _break_cycles(parent_of: dict) -> dict:
    clean = {}
    for child in parent_of:
        seen, cur = {child}, parent_of[child]
        ok = True
        while cur in parent_of:
            if cur in seen:
                ok = False
                break
            seen.add(cur)
            cur = parent_of[cur]
        if ok:
            clean[child] = parent_of[child]
    return clean


def _layout(roots: list, children: dict) -> tuple[dict, int, int]:
    """Tidy-ish tree layout: leaves get consecutive slots, parents centred."""
    pos: dict = {}
    counter = [0]
    max_depth = [0]
    def place(n, depth):
        max_depth[0] = max(max_depth[0], depth)
        kids = children.get(n, [])
        if not kids:
            pos[n] = (counter[0], depth)
            counter[0] += 1
            return
        for k in kids:
            place(k, depth + 1)
        xs = [pos[k][0] for k in kids]
        pos[n] = ((min(xs) + max(xs)) / 2, depth)

    for r in roots:
        place(r, 0)
    return pos, counter[0], max_depth[0] + 1


def _box(ax, x, y, lines, kind, faded=False):
    fill, edge = COLORS.get(kind, COLORS["other"])
    ax.add_patch(FancyBboxPatch((x - BOX_W / 2, y - BOX_H / 2), BOX_W, BOX_H,
                                boxstyle="round,pad=0.02,rounding_size=0.12",
                                fc=fill, ec=edge, lw=1.4, alpha=0.55 if faded else 1, zorder=3))
    ax.text(x, y + 0.22, lines[0], ha="center", va="center", fontsize=8.5, fontweight="bold",
            color="#1a1a1a", zorder=4, clip_on=False)
    ax.text(x, y - 0.14, "\n".join(l for l in lines[1:] if l), ha="center", va="center",
            fontsize=6.8, color="#333333", zorder=4, linespacing=1.25)


def _trim(s: str, n: int = 26) -> str:
    return s if len(s) <= n else s[: n - 1] + "…"


def topology_png(site: dict, wan_label: str = "") -> bytes | None:
    nodes, parent_of, meta = build_graph(site)
    if not nodes:
        return None
    parent_of = _break_cycles(parent_of)
    children: dict = defaultdict(list)
    type_rank = {"gateway": 0, "switch": 1, "olt": 2, "ap": 3}
    for c, p in parent_of.items():
        children[p].append(c)
    for p in children:
        children[p].sort(key=lambda m: (type_rank.get(nodes[m]["type"], 9), nodes[m]["name"].lower()))

    roots = sorted([m for m in nodes if m not in parent_of],
                   key=lambda m: (type_rank.get(nodes[m]["type"], 9), nodes[m]["name"].lower()))
    INTERNET = "__internet__"
    gws = [m for m in roots if nodes[m]["type"] == "gateway"]
    if gws:
        children[INTERNET] = gws
        roots = [INTERNET] + [r for r in roots if r not in gws]
        nodes[INTERNET] = {"name": "Internet", "type": "internet", "model": wan_label, "ip": "", "online": True}

    pos, leaves, depth = _layout(roots, children)
    horizontal = leaves > 6  # left-to-right for wide trees
    sx, sy = 3.0, 2.0  # slot spacing (leaf axis, depth axis)
    if horizontal:
        sx, sy = 1.35, 3.3

    def xy(n):
        slot, d = pos[n]
        return (d * sy, -slot * sx) if horizontal else (slot * sx, -d * sy)

    w_units = (depth * sy if horizontal else leaves * sx) + BOX_W
    h_units = (leaves * sx if horizontal else depth * sy) + BOX_H
    fig_w = min(max(w_units * 0.62, 4.5), 16)
    fig_h = min(max(h_units * 0.62, 3), 22)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h), dpi=200)

    # edges (orthogonal elbows)
    for c, p in list(parent_of.items()) + [(g, INTERNET) for g in children.get(INTERNET, [])]:
        if c not in pos or p not in pos:
            continue
        (x1, y1), (x2, y2) = xy(p), xy(c)
        wireless = meta.get(c, {}).get("wireless")
        style = dict(color="#8A94A6" if not wireless else "#B06A00", lw=1.2,
                     ls=(0, (4, 3)) if wireless else "-", zorder=1)
        if horizontal:
            xm = (x1 + BOX_W / 2 + x2 - BOX_W / 2) / 2
            ax.plot([x1 + BOX_W / 2, xm, xm, x2 - BOX_W / 2], [y1, y1, y2, y2], **style)
            lx, ly, ha, va = xm + 0.08, y2 + 0.08, "left", "bottom"
        else:
            ym = (y1 - BOX_H / 2 + y2 + BOX_H / 2) / 2
            ax.plot([x1, x1, x2, x2], [y1 - BOX_H / 2, ym, ym, y2 + BOX_H / 2], **style)
            lx, ly, ha, va = x2 + 0.08, ym - 0.1, "left", "top"
        label = []
        if meta.get(c, {}).get("port"):
            label.append(f"port {meta[c]['port']}")
        if wireless:
            label.append("mesh")
        if label:
            ax.text(lx, ly, " · ".join(label), fontsize=6, color="#555", ha=ha, va=va, zorder=2,
                    bbox=dict(fc="white", ec="none", pad=0.6, alpha=0.85))

    for n, (slot, d) in pos.items():
        info = nodes[n]
        x, y = xy(n)
        kind = info["type"] if info["type"] in COLORS else "other"
        lines = [_trim(info["name"]), _trim(info["model"], 30), info["ip"]]
        _box(ax, x, y, lines, kind, faded=not info.get("online", True))

    xs = [xy(n)[0] for n in pos]
    ys = [xy(n)[1] for n in pos]
    ax.set_xlim(min(xs) - BOX_W / 2 - 0.3, max(xs) + BOX_W / 2 + 0.9)
    ax.set_ylim(min(ys) - BOX_H / 2 - 0.3, max(ys) + BOX_H / 2 + 0.3)
    ax.set_aspect("equal")
    ax.axis("off")
    _legend(ax, ["gateway", "switch", "ap"], extra_dashed="Wireless (mesh) uplink")
    return _save(fig)


def logical_png(site: dict) -> bytes | None:
    """Gateway -> VLAN networks -> SSIDs mapped onto them."""
    nets = [n for n in site.get("lan_networks", []) if n.get("vlan") or n.get("name")]
    ssids = [s for w in site.get("wlan_groups", []) for s in w.get("ssids", [])]
    if not nets and not ssids:
        return None
    nets = sorted(nets, key=lambda n: n.get("vlan") or 0)
    default_vlan = next((n.get("vlan") for n in nets if n.get("primary")), 1)

    ssid_rows: dict = defaultdict(list)
    seen = set()
    for s in ssids:
        key = (s.get("name"), ssid_vlan(s, default_vlan))
        if key in seen:
            continue
        seen.add(key)
        ssid_rows[key[1]].append(s)

    rows = []  # (network or None, [ssids])
    for n in nets:
        rows.append((n, ssid_rows.pop(n.get("vlan"), [])))
    for v, ss in ssid_rows.items():  # SSIDs on VLANs not defined as networks
        rows.append(({"name": f"VLAN {v}", "vlan": v, "_undefined": True}, ss))

    slots = sum(max(1, len(ss)) for _, ss in rows)
    sy = 1.35
    fig_h = min(max(slots * sy * 0.62 + 0.8, 3), 22)
    fig, ax = plt.subplots(figsize=(10, fig_h), dpi=200)
    gx, nx, sx = 0, 4.2, 8.4
    y = 0.0
    net_ys = []
    for net, ss in rows:
        span = max(1, len(ss))
        ny = y - (span - 1) * sy / 2
        net_ys.append(ny)
        subnet = net.get("gatewaySubnet") or ""
        dh = (net.get("dhcpSettingsVO") or {})
        dhcp = "DHCP on" if dh.get("enable") else ("DHCP off" if dh else "")
        _box(ax, nx, ny, [_trim(f"VLAN {net.get('vlan')} · {net.get('name')}", 30), subnet, dhcp],
             "network", faded=net.get("_undefined", False))
        for i, s in enumerate(ss):
            syy = y - i * sy
            _box(ax, sx, syy, [_trim(s.get("name", ""), 28), security_text(s),
                               band_text(s.get("band"))], "ssid")
            ax.plot([nx + BOX_W / 2, sx - BOX_W / 2], [ny, syy], color="#B06A00", lw=1, zorder=1)
        y -= span * sy
    gy = (net_ys[0] + net_ys[-1]) / 2 if net_ys else 0
    _box(ax, gx, gy, ["Gateway / L3", "Inter-VLAN routing", ""], "gateway")
    for ny in net_ys:
        ax.plot([gx + BOX_W / 2, nx - BOX_W / 2], [gy, ny], color="#2C6CB0", lw=1, zorder=1)
    ax.text(nx, 0.95, "Wired networks (VLANs)", ha="center", fontsize=9, fontweight="bold")
    ax.text(sx, 0.95, "Wireless SSIDs", ha="center", fontsize=9, fontweight="bold")
    ax.set_xlim(gx - BOX_W / 2 - 0.2, sx + BOX_W / 2 + 0.2)
    ax.set_ylim(y + sy - BOX_H / 2 - 0.3, 1.3)
    ax.set_aspect("equal")
    ax.axis("off")
    return _save(fig)


# --------------------------------------------------------------- helpers
SECURITY = {0: "Open", 2: "WPA-Enterprise", 3: "WPA-Personal", 4: "PPSK (no RADIUS)", 5: "PPSK + RADIUS"}
PSK_VERSION = {1: "WPA-PSK", 2: "WPA2-PSK", 3: "WPA/WPA2-PSK", 4: "WPA3-SAE"}
ENT_VERSION = {1: "WPA", 2: "WPA2", 3: "WPA/WPA2", 4: "WPA3"}


def security_text(s: dict) -> str:
    sec = s.get("security")
    if sec == 3 and (s.get("pskSetting") or {}).get("versionPsk") in PSK_VERSION:
        return PSK_VERSION[s["pskSetting"]["versionPsk"]]
    if sec == 2 and (s.get("entSetting") or {}).get("versionEnt") in ENT_VERSION:
        return ENT_VERSION[s["entSetting"]["versionEnt"]] + "-Enterprise"
    if sec == 0 and s.get("oweEnable"):
        return "Enhanced Open (OWE)"
    return SECURITY.get(sec, str(sec) if sec is not None else "")


def band_text(band) -> str:
    if band is None:
        return ""
    parts = [n for bit, n in ((1, "2.4"), (2, "5"), (4, "6")) if int(band) & bit]
    return "/".join(parts) + " GHz" if parts else ""


def ssid_vlan(s: dict, default_vlan=1):
    if s.get("vlanEnable") and s.get("vlanId"):
        return s["vlanId"]
    vs = s.get("vlanSetting") or {}
    if vs.get("mode") == 1:
        cc = vs.get("customConfig") or {}
        m = cc.get("lanNetworkVlanIds") or {}
        if isinstance(m, dict) and m:
            return next(iter(m.values()))
        pool = cc.get("vlanPoolIds") or s.get("vlanPoolIds")
        if pool:
            return pool
    return default_vlan


def _legend(ax, kinds, extra_dashed=None):
    from matplotlib.lines import Line2D
    from matplotlib.patches import Patch
    h = [Patch(fc=COLORS[k][0], ec=COLORS[k][1], label=TYPE_LABEL.get(k, k)) for k in kinds]
    h.append(Patch(fc="white", ec="#999", alpha=0.55, label="Offline / disconnected"))
    if extra_dashed:
        h.append(Line2D([0], [0], color="#B06A00", ls=(0, (4, 3)), label=extra_dashed))
    ax.legend(handles=h, loc="lower center", bbox_to_anchor=(0.5, -0.02), ncol=len(h),
              fontsize=6.5, frameon=False, bbox_transform=ax.figure.transFigure)


def _save(fig) -> bytes:
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=200, bbox_inches="tight", pad_inches=0.15, facecolor="white")
    plt.close(fig)
    return buf.getvalue()
