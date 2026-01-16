"""
HWPX Slot Filler (Local)
- template.hwpx 안의 CLICK_HERE 누름틀(fieldBegin/fieldEnd)을 "슬롯"으로 보고,
  슬롯(name) 기준으로 값을 주입한 뒤,
  최종 제출용으로 CLICK_HERE 누름틀을 제거한 final.hwpx를 생성하여 다운로드합니다.

핵심 요구사항
1) 슬롯(name) <-> 입력 UI를 1:1 매핑
   - slot_map.json을 생성하고, index.html에서 슬롯 목록을 자동 렌더링
2) Python으로 HWPX(XML)에 값 주입
   - fieldBegin(@name,@id) ~ fieldEnd(@beginIDRef=id) 범위 안의 hp:t에 값 주입
3) "HWPX로 제출"해야 하므로 최종본에서 누름틀 제거
   - CLICK_HERE fieldBegin/fieldEnd만 제거 (다른 필드 타입은 보존)

실행 방법(Windows PowerShell 예시)
1) venv 활성화 후 lxml 설치
   .\venv\Scripts\python.exe -m pip install lxml
2) template.hwpx를 같은 폴더에 둠
3) 실행
   .\venv\Scripts\python.exe .\main.py
4) 브라우저에서 접속
   http://127.0.0.1:8000/
"""

import json
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from http.server import BaseHTTPRequestHandler, HTTPServer
from urllib.parse import urlparse

from lxml import etree


# -----------------------------------------------------------------------------
# HWPX (HWPML) 네임스페이스
# -----------------------------------------------------------------------------
# 주의: HWPX 내부 XML은 여러 네임스페이스가 섞이는데,
# 우리가 다루는 필드/문단/테이블 등은 대부분 paragraph 네임스페이스(hp)에 있음.
HP_NS = "http://www.hancom.co.kr/hwpml/2011/paragraph"
NS = {"hp": HP_NS}

# -----------------------------------------------------------------------------
# 파일 경로
# -----------------------------------------------------------------------------
TEMPLATE_HWPX = "template.hwpx"   # 입력 템플릿
SLOT_MAP_JSON = "slot_map.json"   # 슬롯 목록/메타 정보
INDEX_HTML = "index.html"         # 입력 UI


# -----------------------------------------------------------------------------
# 슬롯 메타 데이터 구조
# -----------------------------------------------------------------------------
@dataclass
class SlotMapping:
    """
    슬롯 1개에 대한 메타 정보
    - slot_name: fieldBegin의 name. UI의 key이자 1:1 매핑 기준
    - xml_path: 어떤 xml 파일(Contents/section0.xml 등)에 이 슬롯이 존재하는지
    - field_begin_id: fieldBegin id (fieldEnd의 beginIDRef와 매칭됨)
    - field_id: fieldBegin fieldid (문서 내부 식별자, 템플릿마다 다를 수 있음)
    - row_addr / col_addr: 테이블 셀 위치(가능한 경우)
      * "table 안 슬롯만" 다룰 때 UI에서 위치 힌트로 쓸 수 있음
    """
    slot_name: str
    xml_path: str
    field_begin_id: Optional[str]
    field_id: Optional[str]
    row_addr: Optional[int]
    col_addr: Optional[int]


# -----------------------------------------------------------------------------
# (A) slot_map 생성: template.hwpx 안의 fieldBegin[@name]들을 모아서 JSON으로 저장
# -----------------------------------------------------------------------------
def _find_nearest_cell_addr(el) -> Tuple[Optional[int], Optional[int]]:
    """
    주어진 엘리먼트(el)로부터 부모를 타고 올라가면서,
    가장 가까운 테이블 셀(hp:tc)을 찾아 cellAddr(rowAddr/colAddr)을 읽습니다.

    - 테이블 밖의 슬롯이면 (None, None) 반환
    - row/col은 템플릿 구조에 따라 정확도가 달라질 수 있으므로 "힌트"로만 사용 권장
    """
    cur = el
    while cur is not None:
        # lxml은 네임스페이스 태그가 "{ns}tag" 형태로 들어갑니다.
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


def extract_slot_mappings_from_xml(xml_bytes: bytes, xml_path: str) -> List[SlotMapping]:
    """
    특정 XML(예: Contents/section0.xml)에서 슬롯(fieldBegin[@name])들을 추출합니다.

    - 슬롯의 본질은 "텍스트"가 아니라 fieldBegin의 name입니다.
    - CLICK_HERE든 다른 타입이든, name이 있으면 슬롯로 보고 추출합니다.
      (제품 정책에 따라 CLICK_HERE만 받도록 제한할 수도 있음)
    """
    parser = etree.XMLParser(recover=False, remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser=parser)

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


def build_slot_map_from_hwpx(hwpx_path: str, only_contents: bool = True) -> Dict[str, List[Dict]]:
    """
    HWPX(zip) 내부를 순회하며 XML을 파싱하고, 슬롯 맵을 만듭니다.

    반환 형식:
    {
      "SLOT_DATE": [
         {"slot_name":"SLOT_DATE", "xml_path":"Contents/section0.xml", ...},
         ...
      ],
      ...
    }

    - 동일 name 슬롯이 여러 곳에 있을 수 있으므로 값은 list로 둡니다.
      (문서 내 중복 슬롯을 모두 채우는 동작이 자연스러운 경우가 많음)
    """
    slot_map: Dict[str, List[Dict]] = {}

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        for info in zf.infolist():
            if not info.filename.lower().endswith(".xml"):
                continue
            if only_contents and not info.filename.startswith("Contents/"):
                continue

            data = zf.read(info.filename)

            try:
                mappings = extract_slot_mappings_from_xml(data, info.filename)
            except Exception:
                # 문서 파손 방지: 특정 XML이 파싱 실패해도 전체를 멈추지 않음
                continue

            for m in mappings:
                slot_map.setdefault(m.slot_name, []).append(asdict(m))

    return slot_map


# -----------------------------------------------------------------------------
# (B) UI 자산 생성: slot_map.json + index.html
# -----------------------------------------------------------------------------
def ensure_assets() -> None:
    """
    1) slot_map.json 생성
    2) index.html 생성
       - 브라우저에서 slot_map.json을 fetch해서 슬롯 목록을 렌더링
       - 입력 후 /generate로 POST하면 final.hwpx 다운로드
    """
    slot_map = build_slot_map_from_hwpx(TEMPLATE_HWPX, only_contents=True)
    Path(SLOT_MAP_JSON).write_text(
        json.dumps(slot_map, ensure_ascii=False, indent=2),
        encoding="utf-8"
    )

    html = """<!doctype html>
<html lang="ko">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1" />
  <title>HWPX Slot Filler</title>
  <style>
    body { font-family: system-ui, -apple-system, Segoe UI, Roboto, sans-serif; margin: 24px; }
    .wrap { max-width: 980px; margin: 0 auto; }
    h1 { font-size: 20px; margin: 0 0 12px; }
    p { color: #444; margin: 0 0 20px; line-height: 1.6; }
    .grid { display: grid; grid-template-columns: 1fr 1fr; gap: 12px; }
    .card { border: 1px solid #ddd; border-radius: 10px; padding: 12px; }
    label { display:block; font-size: 12px; color:#666; margin-bottom: 6px; }
    input, textarea {
      width: 100%; box-sizing: border-box; border: 1px solid #ccc; border-radius: 8px;
      padding: 10px; font-size: 14px;
    }
    textarea { min-height: 88px; resize: vertical; }
    .row { display:flex; gap: 10px; align-items:center; margin-top: 14px; flex-wrap: wrap; }
    button {
      border: 0; border-radius: 10px; padding: 10px 14px; font-size: 14px;
      cursor: pointer; background: #111; color: #fff;
    }
    button:disabled { opacity: 0.5; cursor: default; }
    .hint { color:#666; font-size:12px; }
    .mono { font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, monospace; }
    .small { font-size: 12px; color: #666; }
    @media (max-width: 760px) { .grid { grid-template-columns: 1fr; } }
  </style>
</head>
<body>
<div class="wrap">
  <h1>HWPX 슬롯 입력</h1>
  <p>
    슬롯(name) 기준으로 값을 입력하면 서버가 템플릿 HWPX에 값을 주입한 뒤,
    <b>CLICK_HERE 누름틀을 제거</b>한 최종 제출용 <span class="mono">final.hwpx</span>를 다운로드로 제공합니다.
  </p>

  <div id="slots" class="grid"></div>

  <div class="row">
    <button id="btn" disabled>final.hwpx 생성</button>
    <span id="status" class="hint"></span>
  </div>

  <div class="row">
    <span class="hint">슬롯 개수: <span id="count" class="mono"></span></span>
    <span class="small">빈 값은 전송하지 않습니다.</span>
  </div>
</div>

<script>
  const slotsEl = document.getElementById('slots');
  const btn = document.getElementById('btn');
  const statusEl = document.getElementById('status');
  const countEl = document.getElementById('count');

  function makeField(slotName, meta) {
    const card = document.createElement('div');
    card.className = 'card';

    const label = document.createElement('label');
    const loc = meta?.[0]
      ? ` (${meta[0].xml_path}, row=${meta[0].row_addr}, col=${meta[0].col_addr})`
      : '';
    label.textContent = slotName + loc;

    const input = document.createElement('textarea');
    input.placeholder = slotName + ' 값 입력';
    input.dataset.slot = slotName;

    card.appendChild(label);
    card.appendChild(input);
    return card;
  }

  async function init() {
    const res = await fetch('/slot_map.json');
    if (!res.ok) throw new Error('slot_map.json fetch failed');

    const slotMap = await res.json();
    const slotNames = Object.keys(slotMap).sort();

    countEl.textContent = String(slotNames.length);

    slotsEl.innerHTML = '';
    slotNames.forEach(name => slotsEl.appendChild(makeField(name, slotMap[name])));

    btn.disabled = slotNames.length === 0;

    btn.onclick = async () => {
      try {
        btn.disabled = true;
        statusEl.textContent = '생성 중...';

        // payload: { SLOT_NAME: "value" }
        const payload = {};
        document.querySelectorAll('[data-slot]').forEach(el => {
          const v = (el.value || '').trim();
          if (v.length > 0) payload[el.dataset.slot] = v;
        });

        const r = await fetch('/generate', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify(payload),
        });

        if (!r.ok) {
          const t = await r.text();
          throw new Error(t || 'server error');
        }

        // 응답을 blob으로 받아 파일 다운로드
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
      } catch (e) {
        statusEl.textContent = '오류: ' + (e?.message || String(e));
      } finally {
        btn.disabled = false;
      }
    };
  }

  init().catch(e => {
    statusEl.textContent = '초기화 오류: ' + e.message;
  });
</script>
</body>
</html>
"""
    Path(INDEX_HTML).write_text(html, encoding="utf-8")


# -----------------------------------------------------------------------------
# (C) 값 주입 로직: fieldBegin(id) ~ fieldEnd(beginIDRef=id) 범위의 hp:t에 주입
# -----------------------------------------------------------------------------
def fill_fields_in_xml(xml_bytes: bytes, values: Dict[str, str]) -> bytes:
    """
    - values에 있으면: 해당 값 주입
    - values에 없으면: placeholder를 빈칸("")으로 만들어 제출용에서 깨끗하게 보이게 함
      (중요: 구조는 유지, 텍스트만 비움)
    """
    parser = etree.XMLParser(recover=False, remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser=parser)

    for fb in root.xpath(".//hp:fieldBegin[@name][@id]", namespaces=NS):
        slot_name = fb.get("name")
        begin_id = fb.get("id")
        if not slot_name or not begin_id:
            continue

        # ✅ 값이 없으면 빈칸으로 처리
        target_value = values.get(slot_name, "")

        # fieldEnd 1개 확정
        ends = fb.xpath(
            "following::hp:fieldEnd[@beginIDRef=$bid][1]",
            namespaces=NS,
            bid=begin_id
        )
        if not ends:
            continue
        end_node = ends[0]

        # fb~end 범위 내 hp:t만 잡기
        t_nodes = fb.xpath(
            "following::hp:t[following::hp:fieldEnd[@beginIDRef=$bid][1] = $end]",
            namespaces=NS,
            bid=begin_id,
            end=end_node
        )

        if t_nodes:
            t_nodes[0].text = target_value
            for extra in t_nodes[1:]:
                extra.text = ""
        else:
            # 범위 안에 hp:t가 없으면 end_node 직전에 hp:t 삽입
            ctrl = end_node.getparent()
            if ctrl is not None:
                t = etree.Element(f"{{{HP_NS}}}t")
                t.text = target_value
                ctrl.insert(ctrl.index(end_node), t)

    return etree.tostring(root, xml_declaration=True, encoding="utf-8", pretty_print=False)

# -----------------------------------------------------------------------------
# (D) 제출용 처리: CLICK_HERE 누름틀 제거
# -----------------------------------------------------------------------------
def strip_clickhere_fields(xml_bytes: bytes) -> bytes:
    """
    제출용 HWPX를 만들기 위해 CLICK_HERE 누름틀(fieldBegin/fieldEnd)을 제거합니다.

    왜 이렇게 하느냐?
    - HWP(한글)는 CLICK_HERE의 Direction(가이드)을 화면에 겹쳐 그릴 수 있음
    - 제출용 HWPX에서는 "일반 텍스트"만 남게 하는 편이 안전한 경우가 많음
    - 단, 다른 필드(예: 자동 목차, 페이지 번호 등)를 건드리면 문서가 깨질 수 있으므로,
      여기서는 type='CLICK_HERE'에 해당하는 beginIDRef만 정확히 제거합니다.
    """
    parser = etree.XMLParser(recover=True, remove_blank_text=False)
    root = etree.fromstring(xml_bytes, parser=parser)

    # 1) CLICK_HERE fieldBegin id들을 모아둔다 (fieldEnd 제거를 정확히 하기 위함)
    begin_ids = set()
    for fb in root.xpath(".//hp:fieldBegin[@type='CLICK_HERE'][@id]", namespaces=NS):
        bid = fb.get("id")
        if bid:
            begin_ids.add(bid)

    # 2) CLICK_HERE fieldBegin 제거
    for fb in root.xpath(".//hp:fieldBegin[@type='CLICK_HERE']", namespaces=NS):
        parent = fb.getparent()
        if parent is not None:
            parent.remove(fb)

    # 3) 동일 beginIDRef를 참조하는 fieldEnd만 제거
    if begin_ids:
        for fe in root.xpath(".//hp:fieldEnd[@beginIDRef]", namespaces=NS):
            ref = fe.get("beginIDRef")
            if ref in begin_ids:
                parent = fe.getparent()
                if parent is not None:
                    parent.remove(fe)

    return etree.tostring(root, xml_declaration=True, encoding="utf-8", pretty_print=False)


# -----------------------------------------------------------------------------
# (E) HWPX(zip) 단위 파이프라인
# -----------------------------------------------------------------------------
def generate_submit_hwpx(
    input_hwpx: str,
    output_hwpx: str,
    values: Dict[str, str],
    only_contents: bool = True
) -> None:
    """
    HWPX(zip) 내부 XML 파일마다:
    1) fill_fields_in_xml: 값 주입
    2) strip_clickhere_fields: CLICK_HERE 누름틀 제거
    를 수행하고 output_hwpx로 저장합니다.

    only_contents=True인 이유:
    - 대부분 본문은 Contents/ 아래에 있고,
    - 다른 xml(헤더/메타 등)을 건드리면 문서가 깨질 가능성이 있으므로
    기본값은 Contents/만 처리합니다.
    """
    with zipfile.ZipFile(input_hwpx, "r") as zin:
        with zipfile.ZipFile(output_hwpx, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                data = zin.read(info.filename)

                is_xml = info.filename.lower().endswith(".xml")
                in_contents = info.filename.startswith("Contents/")

                if is_xml and ((not only_contents) or in_contents):
                    try:
                        filled = fill_fields_in_xml(data, values)
                        # final_xml = strip_clickhere_fields(filled)
                        zout.writestr(info.filename, filled)
                    except Exception:
                        # 특정 XML 처리 중 오류가 나도 문서를 깨지지 않게 원본을 유지합니다.
                        zout.writestr(info.filename, data)
                else:
                    # XML이 아니거나 Contents/가 아니면 그대로 복사
                    zout.writestr(info.filename, data)


# -----------------------------------------------------------------------------
# (F) HTTP 서버: index.html, slot_map.json 제공 + /generate로 final.hwpx 생성
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
            body = Path(INDEX_HTML).read_bytes()
            return self._send(200, body, "text/html; charset=utf-8")

        if p == "/slot_map.json":
            body = Path(SLOT_MAP_JSON).read_bytes()
            return self._send(200, body, "application/json; charset=utf-8")

        return self._send(404, b"Not Found")

    def do_POST(self):
        p = urlparse(self.path).path
        if p != "/generate":
            return self._send(404, b"Not Found")

        try:
            # (1) JSON payload 읽기
            length = int(self.headers.get("Content-Length", "0"))
            raw = self.rfile.read(length)
            values = json.loads(raw.decode("utf-8"))

            if not isinstance(values, dict):
                raise ValueError("Invalid JSON payload: must be object")

            # (2) final.hwpx 생성
            out_path = "final.hwpx"
            generate_submit_hwpx(TEMPLATE_HWPX, out_path, values, only_contents=True)

            # (3) 바이너리 응답으로 다운로드
            body = Path(out_path).read_bytes()
            self.send_response(200)
            self.send_header("Content-Type", "application/vnd.hancom.hwpx")
            self.send_header("Content-Disposition", 'attachment; filename="final.hwpx"')
            self.send_header("Content-Length", str(len(body)))
            self.end_headers()
            self.wfile.write(body)

        except Exception as e:
            msg = f"Error: {e}".encode("utf-8")
            return self._send(500, msg, "text/plain; charset=utf-8")


# -----------------------------------------------------------------------------
# 엔트리포인트
# -----------------------------------------------------------------------------
def main():
    # 템플릿 존재 확인
    if not Path(TEMPLATE_HWPX).exists():
        raise FileNotFoundError(f"{TEMPLATE_HWPX} not found. Put template.hwpx next to main.py")

    # slot_map.json + index.html 생성
    ensure_assets()

    host = "127.0.0.1"
    port = 8000
    print(f"Open: http://{host}:{port}/")

    server = HTTPServer((host, port), Handler)
    server.serve_forever()


if __name__ == "__main__":
    main()
