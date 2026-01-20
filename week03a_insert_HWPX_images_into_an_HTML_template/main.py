# main.py
"""
HWPX Slot Filler (Local) - Safe Edition
- 텍스트 슬롯: CLICK_HERE fieldBegin/fieldEnd 범위 안의 hp:t 텍스트 치환
- 이미지 슬롯(IMG_ 접두사): 해당 슬롯의 fieldEnd "뒤"에서 가장 가까운 hp:pic을 찾아,
  그 안의 hc:img@binaryItemIDRef(예: image1)가 가리키는 BinData 파일을 교체

중요: 한글이 HWPX 구조에 매우 민감하므로
- XML 재직렬화를 최소화
- 파싱 실패 시 원본 유지
- (기본값) CLICK_HERE 제거는 OFF 로 둠 (열림 문제 방지)

실행:
  .\venv\Scripts\python.exe -m pip install lxml
  .\venv\Scripts\python.exe .\main.py
  http://127.0.0.1:8000/
"""

import base64
import json
import re
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any

from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse

from lxml import etree


# -----------------------------------------------------------------------------
# Config
# -----------------------------------------------------------------------------
TEMPLATE_HWPX = "template.hwpx"
SLOT_MAP_JSON = "slot_map.json"
INDEX_HTML = "index.html"

# 제출용 CLICK_HERE 제거: 파일 열림 이슈가 있으면 False로 유지하세요.
STRIP_CLICK_HERE = False

# 이미지 슬롯 prefix
IMG_PREFIX = "IMG_"

# -----------------------------------------------------------------------------
# Namespaces
# -----------------------------------------------------------------------------
HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
HC_NS = "http://www.hancom.co.kr/hwpml/2011/core"
NS = {"hp": HP_NS, "hc": HC_NS}


# -----------------------------------------------------------------------------
# Slot meta
# -----------------------------------------------------------------------------
@dataclass
class SlotMapping:
    slot_name: str
    xml_path: str
    field_begin_id: Optional[str]
    field_id: Optional[str]
    row_addr: Optional[int]
    col_addr: Optional[int]


# -----------------------------------------------------------------------------
# Utils
# -----------------------------------------------------------------------------
def _xml_parse_strict(xml_bytes: bytes) -> etree._Element:
    # recover=False: 한글 호환성 위해 "임의 복구 후 저장"을 피함
    parser = etree.XMLParser(recover=False, remove_blank_text=False, huge_tree=True)
    return etree.fromstring(xml_bytes, parser=parser)


def _xml_dump(root: etree._Element) -> bytes:
    # pretty_print=False로 구조 변경 최소화
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", pretty_print=False)


def _find_nearest_cell_addr(el) -> Tuple[Optional[int], Optional[int]]:
    cur = el
    while cur is not None:
        if cur.tag == f"{{{HP_NS}}}tc":
            cell_addr = cur.find("hp:cellAddr", namespaces=NS)
            if cell_addr is None:
                return None, None
            try:
                row = int(cell_addr.get("rowAddr")) if cell_addr.get("rowAddr") is not None else None
            except ValueError:
                row = None
            try:
                col = int(cell_addr.get("colAddr")) if cell_addr.get("colAddr") is not None else None
            except ValueError:
                col = None
            return row, col
        cur = cur.getparent()
    return None, None


def decode_data_url(data_url: str) -> Tuple[str, bytes]:
    if not data_url.startswith("data:"):
        raise ValueError("Invalid dataUrl (must start with data:)")
    header, b64 = data_url.split(",", 1)
    m = re.match(r"data:([^;]+);base64", header)
    if not m:
        raise ValueError("Invalid dataUrl header")
    mime = m.group(1).strip().lower()
    raw = base64.b64decode(b64)
    return mime, raw


def is_xml_path(name: str) -> bool:
    return name.lower().endswith(".xml")


# -----------------------------------------------------------------------------
# (A) slot_map 생성
# -----------------------------------------------------------------------------
def extract_slot_mappings_from_xml(xml_bytes: bytes, xml_path: str) -> List[SlotMapping]:
    root = _xml_parse_strict(xml_bytes)

    out: List[SlotMapping] = []
    for fb in root.xpath(".//hp:fieldBegin[@name]", namespaces=NS):
        slot_name = fb.get("name")
        if not slot_name:
            continue
        row_addr, col_addr = _find_nearest_cell_addr(fb)
        out.append(
            SlotMapping(
                slot_name=slot_name,
                xml_path=xml_path,
                field_begin_id=fb.get("id"),
                field_id=fb.get("fieldid"),
                row_addr=row_addr,
                col_addr=col_addr,
            )
        )
    return out


def build_slot_map_from_hwpx(hwpx_path: str, only_contents: bool = True) -> Dict[str, List[Dict[str, Any]]]:
    slot_map: Dict[str, List[Dict[str, Any]]] = {}

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        for info in zf.infolist():
            if not is_xml_path(info.filename):
                continue
            if only_contents and not info.filename.startswith("Contents/"):
                continue

            data = zf.read(info.filename)
            try:
                mappings = extract_slot_mappings_from_xml(data, info.filename)
            except Exception:
                continue

            for m in mappings:
                slot_map.setdefault(m.slot_name, []).append(asdict(m))

    return slot_map


# -----------------------------------------------------------------------------
# (B) index.html 생성 (IMG_는 file 업로드)
# -----------------------------------------------------------------------------
def ensure_assets() -> None:
    slot_map = build_slot_map_from_hwpx(TEMPLATE_HWPX, only_contents=True)
    Path(SLOT_MAP_JSON).write_text(json.dumps(slot_map, ensure_ascii=False, indent=2), encoding="utf-8")

    html = f"""<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>HWPX Slot Filler</title>
  <style>
    body {{ font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; margin: 24px; }}
    .wrap {{ max-width: 980px; margin: 0 auto; }}
    h1 {{ font-size: 20px; margin: 0 0 12px; }}
    p {{ color: #444; margin: 0 0 20px; line-height: 1.6; }}
    .grid {{ display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }}
    .card {{ border: 1px solid #ddd; border-radius: 10px; padding: 12px; }}
    label {{ display:block; font-size: 12px; color:#666; margin-bottom: 6px; }}
    input, textarea {{
      width: 100%; box-sizing: border-box; border: 1px solid #ccc; border-radius: 8px;
      padding: 10px; font-size: 14px;
    }}
    input[type="file"] {{ padding: 8px; }}
    textarea {{ min-height: 88px; resize: vertical; }}
    .row {{ display:flex; gap: 10px; align-items:center; margin-top: 14px; flex-wrap: wrap; }}
    button {{
      border: 0; border-radius: 10px; padding: 10px 14px; font-size: 14px;
      cursor: pointer; background: #111; color: #fff;
    }}
    button:disabled {{ opacity: 0.5; cursor: default; }}
    .hint {{ color:#666; font-size:12px; }}
    .mono {{ font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }}
    .small {{ font-size: 12px; color: #666; }}
    .preview {{ margin-top: 8px; font-size: 12px; color: #666; }}
    @media (max-width: 760px) {{ .grid {{ grid-template-columns: 1fr; }} }}
  </style>
</head>
<body>
<div class="wrap">
  <h1>HWPX 슬롯 입력</h1>
  <p>
    텍스트 슬롯은 값이 주입되고, 이미지 슬롯(<span class="mono">{IMG_PREFIX}</span>로 시작)은 파일 업로드로 처리됩니다.
    <span class="small">(STRIP_CLICK_HERE={str(STRIP_CLICK_HERE).lower()})</span>
  </p>

  <div id="slots" class="grid"></div>

  <div class="row">
    <button id="btn" disabled>final.hwpx 생성</button>
    <span id="status" class="hint"></span>
  </div>

  <div class="row">
    <span class="hint">슬롯 개수: <span id="count" class="mono"></span></span>
    <span class="small">텍스트 빈 값은 전송하지 않습니다. 이미지 슬롯은 파일 선택 시만 전송합니다.</span>
  </div>
</div>

<script>
  const slotsEl = document.getElementById('slots');
  const btn = document.getElementById('btn');
  const statusEl = document.getElementById('status');
  const countEl = document.getElementById('count');

  const imageState = {{}};

  function makeTextField(slotName, meta) {{
    const card = document.createElement('div');
    card.className = 'card';

    const label = document.createElement('label');
    const loc = meta?.[0]
      ? ` ({{${{meta[0].xml_path}}}}, row={{${{meta[0].row_addr}}}}, col={{${{meta[0].col_addr}}}})`
      : '';
    label.textContent = slotName + loc;

    const input = document.createElement('textarea');
    input.placeholder = slotName + ' 값 입력';
    input.dataset.slot = slotName;

    card.appendChild(label);
    card.appendChild(input);
    return card;
  }}

  function makeImageField(slotName, meta) {{
    const card = document.createElement('div');
    card.className = 'card';

    const label = document.createElement('label');
    const loc = meta?.[0]
      ? ` ({{${{meta[0].xml_path}}}}, row={{${{meta[0].row_addr}}}}, col={{${{meta[0].col_addr}}}})`
      : '';
    label.textContent = slotName + loc + " (이미지 업로드)";

    const input = document.createElement('input');
    input.type = 'file';
    input.accept = 'image/*';
    input.dataset.imgslot = slotName;

    const preview = document.createElement('div');
    preview.className = 'preview';
    preview.textContent = '선택된 파일 없음';

    input.onchange = async () => {{
      const f = input.files && input.files[0];
      if (!f) {{
        delete imageState[slotName];
        preview.textContent = '선택된 파일 없음';
        return;
      }}
      const dataUrl = await new Promise((resolve, reject) => {{
        const r = new FileReader();
        r.onload = () => resolve(r.result);
        r.onerror = () => reject(new Error('file read error'));
        r.readAsDataURL(f);
      }});
      imageState[slotName] = {{ filename: f.name, dataUrl }};
      preview.textContent = `선택됨: ${{f.name}} (${{Math.round(f.size/1024)}}KB)`;
    }};

    card.appendChild(label);
    card.appendChild(input);
    card.appendChild(preview);
    return card;
  }}

  async function init() {{
    const res = await fetch('/slot_map.json');
    if (!res.ok) throw new Error('slot_map.json fetch failed');

    const slotMap = await res.json();
    const slotNames = Object.keys(slotMap).sort();

    countEl.textContent = String(slotNames.length);
    slotsEl.innerHTML = '';

    slotNames.forEach(name => {{
      if (name.startsWith('{IMG_PREFIX}')) {{
        slotsEl.appendChild(makeImageField(name, slotMap[name]));
      }} else {{
        slotsEl.appendChild(makeTextField(name, slotMap[name]));
      }}
    }});

    btn.disabled = slotNames.length === 0;

    btn.onclick = async () => {{
      try {{
        btn.disabled = true;
        statusEl.textContent = '생성 중...';

        const payload = {{}};
        document.querySelectorAll('[data-slot]').forEach(el => {{
          const v = (el.value || '').trim();
          if (v.length > 0) payload[el.dataset.slot] = v;
        }});

        Object.keys(imageState).forEach(k => {{
          payload[k] = imageState[k];
        }});

        const r = await fetch('/generate', {{
          method: 'POST',
          headers: {{ 'Content-Type': 'application/json' }},
          body: JSON.stringify(payload),
        }});

        if (!r.ok) {{
          const t = await r.text();
          throw new Error(t || 'server error');
        }}

        const blob = await r.blob();
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = 'final.hwpx';
        document.body.appendChild(a);
        a.click();
        a.remove();
        URL.revokeObjectURL(url);

        statusEl.textContent = '완료: final.hwpx 다운로드됨';
      }} catch (e) {{
        statusEl.textContent = '오류: ' + (e?.message || String(e));
      }} finally {{
        btn.disabled = false;
      }}
    }};
  }}

  init().catch(e => {{
    statusEl.textContent = '초기화 오류: ' + e.message;
  }});
</script>
</body>
</html>
"""
    Path(INDEX_HTML).write_text(html, encoding="utf-8")


# -----------------------------------------------------------------------------
# (C) 텍스트 주입: fieldBegin~fieldEnd 범위의 hp:t 치환
# -----------------------------------------------------------------------------
def fill_fields_in_xml(xml_bytes: bytes, values: Dict[str, str]) -> Tuple[bytes, bool]:
    """
    returns: (new_bytes, changed)
    """
    root = _xml_parse_strict(xml_bytes)
    changed = False

    for fb in root.xpath(".//hp:fieldBegin[@name][@id]", namespaces=NS):
        slot_name = fb.get("name")
        begin_id = fb.get("id")
        if not slot_name or not begin_id:
            continue
        if slot_name not in values:
            continue
        # 이미지 슬롯은 여기서 제외
        if slot_name.startswith(IMG_PREFIX):
            continue

        target_value = values[slot_name]

        ends = fb.xpath(
            "following::hp:fieldEnd[@beginIDRef=$bid][1]",
            namespaces=NS,
            bid=begin_id
        )
        if not ends:
            continue
        end_node = ends[0]

        t_nodes = fb.xpath(
            "following::hp:t[following::hp:fieldEnd[@beginIDRef=$bid][1] = $end]",
            namespaces=NS,
            bid=begin_id,
            end=end_node
        )

        if t_nodes:
            if (t_nodes[0].text or "") != target_value:
                t_nodes[0].text = target_value
                changed = True
            for extra in t_nodes[1:]:
                if extra.text:
                    extra.text = ""
                    changed = True

    if not changed:
        return xml_bytes, False
    return _xml_dump(root), True


# -----------------------------------------------------------------------------
# (D) CLICK_HERE 제거 (기본 OFF)
# -----------------------------------------------------------------------------
def strip_clickhere_fields(xml_bytes: bytes) -> Tuple[bytes, bool]:
    root = _xml_parse_strict(xml_bytes)
    changed = False

    begin_ids = set()
    for fb in root.xpath(".//hp:fieldBegin[@type='CLICK_HERE'][@id]", namespaces=NS):
        bid = fb.get("id")
        if bid:
            begin_ids.add(bid)

    # fieldBegin 제거
    fbs = root.xpath(".//hp:fieldBegin[@type='CLICK_HERE']", namespaces=NS)
    for fb in fbs:
        parent = fb.getparent()
        if parent is not None:
            parent.remove(fb)
            changed = True

    # fieldEnd 제거 (beginIDRef 일치하는 것만)
    if begin_ids:
        fes = root.xpath(".//hp:fieldEnd[@beginIDRef]", namespaces=NS)
        for fe in fes:
            ref = fe.get("beginIDRef")
            if ref in begin_ids:
                parent = fe.getparent()
                if parent is not None:
                    parent.remove(fe)
                    changed = True

    if not changed:
        return xml_bytes, False
    return _xml_dump(root), True


# -----------------------------------------------------------------------------
# (E) 이미지: IMG 슬롯의 fieldEnd "뒤" 첫 hp:pic -> binaryItemIDRef 찾기
# -----------------------------------------------------------------------------
def find_pic_binary_id_after_field(xml_bytes: bytes, slot_name: str) -> Optional[str]:
    root = _xml_parse_strict(xml_bytes)

    fbs = root.xpath(".//hp:fieldBegin[@name=$nm][@id]", namespaces=NS, nm=slot_name)
    if not fbs:
        return None

    for fb in fbs:
        begin_id = fb.get("id")
        if not begin_id:
            continue

        ends = fb.xpath("following::hp:fieldEnd[@beginIDRef=$bid][1]", namespaces=NS, bid=begin_id)
        if not ends:
            continue
        end_node = ends[0]

        pic = end_node.xpath("following::hp:pic[1]", namespaces=NS)
        if not pic:
            continue
        pic = pic[0]

        img = pic.xpath(".//hc:img[@binaryItemIDRef][1]", namespaces=NS)
        if not img:
            continue
        return img[0].get("binaryItemIDRef")

    return None


# -----------------------------------------------------------------------------
# (F) binaryItemIDRef -> BinData 경로 해석 (content.hpf 포함, 강건 탐색)
# -----------------------------------------------------------------------------
def resolve_bindata_href_for_binary_item(all_xml: Dict[str, bytes], binary_id: str) -> Optional[str]:
    """
    image1 같은 binary_id가 가리키는 BinData/xxx 파일 경로를 찾는다.
    문서마다 관계파일 구조가 달라서, "가능한 많이" 탐색한다.
    """
    # 1) XML DOM 기반: @id=binary_id 이면서 @href에 BinData/ 포함
    for path, xml_bytes in all_xml.items():
        if not is_xml_path(path) and not path.lower().endswith(".hpf"):
            continue

        try:
            root = _xml_parse_strict(xml_bytes)
        except Exception:
            continue

        nodes = root.xpath(
            "//*[@id=$bid and @href and (contains(@href, 'BinData/') or contains(@href, 'BinData\\'))]",
            bid=binary_id
        )
        if nodes:
            href = nodes[0].get("href")
            if href:
                return href.replace("\\", "/")

        # 흔한 대체 키들
        nodes = root.xpath(
            "//*[@itemID=$bid or @itemId=$bid or @binaryItemIDRef=$bid]",
            bid=binary_id
        )
        for n in nodes:
            href = n.get("href")
            if href and ("BinData/" in href or "BinData\\" in href):
                return href.replace("\\", "/")

    # 2) 원문 문자열 기반 fallback (binary_id 주변에서 BinData 경로 추출)
    pat = re.compile(rf'(?s){re.escape(binary_id)}.*?(BinData[\\/][^"\'<\s]+)')
    for path, xml_bytes in all_xml.items():
        if not is_xml_path(path) and not path.lower().endswith(".hpf"):
            continue
        s = xml_bytes.decode("utf-8", errors="ignore")
        m = pat.search(s)
        if m:
            return m.group(1).replace("\\", "/")

    return None


# -----------------------------------------------------------------------------
# (G) HWPX 생성: 최소 변경 복사
# -----------------------------------------------------------------------------
def generate_submit_hwpx(input_hwpx: str, output_hwpx: str, payload: Dict[str, Any], only_contents: bool = True) -> None:
    # 분리
    text_values: Dict[str, str] = {}
    image_values: Dict[str, Dict[str, str]] = {}

    for k, v in (payload or {}).items():
        if isinstance(k, str) and k.startswith(IMG_PREFIX) and isinstance(v, dict) and "dataUrl" in v:
            image_values[k] = v
        elif isinstance(v, (str, int, float)):
            text_values[k] = str(v)

    with zipfile.ZipFile(input_hwpx, "r") as zin:
        # XML/HFP preload (매핑 해석용)
        all_xml: Dict[str, bytes] = {}
        for info in zin.infolist():
            if is_xml_path(info.filename) or info.filename.lower().endswith(".hpf"):
                all_xml[info.filename] = zin.read(info.filename)

        # 이미지 슬롯 매핑: slot -> binaryId -> bindata path
        slot_to_bindata: Dict[str, str] = {}
        for slot_name in image_values.keys():
            binary_id = None
            # contents xml들에서 slot 탐색
            for path, xml_bytes in all_xml.items():
                if only_contents and not path.startswith("Contents/"):
                    continue
                if not is_xml_path(path):
                    continue
                try:
                    bid = find_pic_binary_id_after_field(xml_bytes, slot_name)
                except Exception:
                    continue
                if bid:
                    binary_id = bid
                    break

            if not binary_id:
                continue

            href = resolve_bindata_href_for_binary_item(all_xml, binary_id)
            if href:
                slot_to_bindata[slot_name] = href

        # 업로드 이미지 디코드
        slot_to_imgbytes: Dict[str, bytes] = {}
        for slot_name, info in image_values.items():
            mime, raw = decode_data_url(str(info["dataUrl"]))
            # 한글 HWPX는 보통 png/jpg가 안전
            if not (mime.endswith("png") or mime.endswith("jpeg") or mime.endswith("jpg") or mime.endswith("bmp")):
                raise ValueError(f"{slot_name}: unsupported image mime: {mime}")
            slot_to_imgbytes[slot_name] = raw

        missing = [s for s in image_values.keys() if s not in slot_to_bindata]
        if missing:
            print("[WARN] 이미지 슬롯에서 BinData 매핑을 찾지 못했습니다:", missing)

        # 실제 zip 작성: 원본을 그대로 복사 + 필요한 항목만 변경
        with zipfile.ZipFile(output_hwpx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                name = info.filename
                data = zin.read(name)

                # (1) Contents XML: 슬롯이 있는 파일만 "최소 변경"
                if is_xml_path(name) and ((not only_contents) or name.startswith("Contents/")):
                    # 슬롯명이 포함된 xml만 변경 시도(불필요 재직렬화 방지)
                    # 단순 바이트 검사로 빠르게 판단
                    maybe_related = False
                    for s in text_values.keys():
                        if s.encode("utf-8") in data:
                            maybe_related = True
                            break
                    if not maybe_related:
                        for s in image_values.keys():
                            if s.encode("utf-8") in data:
                                maybe_related = True
                                break

                    if maybe_related:
                        try:
                            new_data, changed1 = fill_fields_in_xml(data, text_values)
                            changed2 = False
                            if STRIP_CLICK_HERE:
                                new_data2, changed2 = strip_clickhere_fields(new_data)
                                new_data = new_data2
                            if changed1 or changed2:
                                zout.writestr(name, new_data)
                            else:
                                zout.writestr(name, data)
                        except Exception:
                            # 파싱/직렬화 실패 시 원본 유지
                            zout.writestr(name, data)
                    else:
                        zout.writestr(name, data)
                    continue

                # (2) BinData overwrite: 매핑된 파일만 교체
                if name.startswith("BinData/"):
                    replaced = False
                    for slot_name, bindata_path in slot_to_bindata.items():
                        if bindata_path == name and slot_name in slot_to_imgbytes:
                            zout.writestr(name, slot_to_imgbytes[slot_name])
                            replaced = True
                            break
                    if replaced:
                        continue

                # (3) 그 외: 원본 그대로 복사
                zout.writestr(name, data)


# -----------------------------------------------------------------------------
# HTTP Server
# -----------------------------------------------------------------------------
class Handler(BaseHTTPRequestHandler):
    def _send(self, code: int, body: bytes, content_type: str = "text/plain; charset=utf-8"):
        self.send_response(code)
        self.send_header("Content-Type", content_type)
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def do_GET(self):
        p = urlparse(self.path).path

        if p == "/" or p == "/index.html":
            return self._send(200, Path(INDEX_HTML).read_bytes(), "text/html; charset=utf-8")

        if p == "/slot_map.json":
            return self._send(200, Path(SLOT_MAP_JSON).read_bytes(), "application/json; charset=utf-8")

        return self._send(404, b"Not Found")

    def do_POST(self):
        p = urlparse(self.path).path
        if p != "/generate":
            return self._send(404, b"Not Found")

        try:
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length)
            payload = json.loads(raw.decode("utf-8"))
            if not isinstance(payload, dict):
                raise ValueError("Invalid JSON payload: must be object")

            out_path = "final.hwpx"
            generate_submit_hwpx(TEMPLATE_HWPX, out_path, payload, only_contents=True)

            body = Path(out_path).read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.hancom.hwpx")
            self.send_header("Content-Disposition", 'attachment; filename="final.hwpx"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        except Exception as e:
            return self._send(500, f"Error: {e}".encode("utf-8"), "text/plain; charset=utf-8")


def main():
    if not Path(TEMPLATE_HWPX).exists():
        raise FileNotFoundError(f"{TEMPLATE_HWPX} not found. Put template.hwpx next to main.py")

    ensure_assets()

    host = "127.0.0.1"
    port = 8000
    print(f"Open: http://{host}:{port}/ (STRIP_CLICK_HERE={STRIP_CLICK_HERE})")

    server = HTTPServer((host, port), Handler)
    server.serve_forever()


if __name__ == "__main__":
    main()
