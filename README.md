# Omada As-Built Generator

Interrogates a self-hosted TP-Link Omada Software Controller (built and checked against the
v6.x Open API; works on 5.12+) and writes a Word as-built document:

- Controller summary and a per-site summary
- Site details, location and floor-plan maps
- Device inventory (model, MAC, IP, serial, firmware, status)
- Physical topology diagram (gateway → switches → APs, uplink ports, mesh links) + uplink table
- Internet / WAN status and connection settings (PPPoE/L2TP credentials)
- Networks & VLANs (subnet, DHCP range, DNS, lease)
- Switch port profiles (native / tagged / voice VLANs) and per-switch port assignments
- WLAN groups and SSIDs **with passphrases**, security, bands, VLAN, 802.11r, PMF
- Access points (WLAN group, uplink, channels)
- Logical diagram (VLANs → SSIDs)
- DHCP reservations
- Appendix listing anything the controller didn't return

## 1. Create an Open API application on the controller

1. Global View → **Settings → Platform Integration → Open API** → **Add New App**.
2. Mode: **Client** (client credentials). Role: a read-only/Viewer role covers most of the
   document. If SSID passphrases or WAN credentials come back blank, the role likely lacks access
   to those settings — give the app an Administrator role (the tool only ever reads).
3. Grant it the sites you want documented.
4. Note the **Client ID**, **Client Secret** and the **Interface Access Address**
   (e.g. `https://controller:8043`). The Omada ID is detected automatically.

The controller's own API reference is linked from the same page (top-right), and its
machine-readable spec is served at `https://<controller>:8043/v3/api-docs`.

## 2. Install

```bash
python -m venv .venv
.venv\Scripts\activate
pip3 install -r requirements.txt
copy omada.ini.example omada.ini   # then edit it
```

## 3. Run

```bash
# check connectivity + credentials, list sites, save the controller's API spec
python3 -m omada_asbuilt --probe

# full as-built for all sites
python3 -m omada_asbuilt --customer "Customer Name" --author "Author Name"

# one or two sites, passwords masked (for a distribution copy)
python3 -m omada_asbuilt --site "Head Office" --site "Warehouse" --secrets mask

# re-render from the saved raw data, no controller access
python3 -m omada_asbuilt --from-json output/omada-raw-2026-09-28.json --secrets omit
```

Output goes to `./output/`: the `.docx` and `omada-raw-<date>.json` (the unedited API
responses). **The raw JSON always contains passwords**, whatever `--secrets` says.

When Word opens the document it asks to update fields — click **Yes** to build the table of
contents and page numbers.

Self-signed certificate: set `insecure = true` in `omada.ini` (or `--insecure`), or
point `ca_bundle` at the controller's certificate.



## Layout

```
omada_asbuilt/
  client.py     Open API (OAuth client-credentials, token refresh, paging) + optional web API
  collector.py  what gets pulled from the controller
  diagrams.py   topology + logical diagrams (matplotlib, no Graphviz needed)
  report.py     the Word document (python-docx)

```