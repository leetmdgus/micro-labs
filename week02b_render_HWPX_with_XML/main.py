from __future__ import annotations

import zipfile
import html
import re
import mimetypes
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from lxml import etree


# -------------------------
# Utils
# -------------------------

def local_name(el: etree._Element) -> str:
    return etree.QName(el).localname

def text_content(node: etree._Element) -> str:
    # HWPX는 텍스트가 여러 태그로 쪼개져 있을 수 있어 itertext로 안전하게 수집
    s = "".join(node.itertext())
    return " ".join(s.split())  # 공백 정리(줄바꿈 보존 원하면 이 줄 제거)

def esc(s: str) -> str:
    return html.escape(s or "")

def read_xml_from_zip(zf: zipfile.ZipFile, path: str) -> Optional[etree._Element]:
    try:
        raw = zf.read(path)
        return etree.fromstring(raw)
    except Exception:
        return None


# -------------------------
# Image helpers (BinData + placeholder)
# -------------------------

IMG_PLACEHOLDER_RE = re.compile(
    r"원본\s*그림의\s*이름:\s*([A-Za-z0-9_.-]+\.(?:bmp|png|jpg|jpeg|gif|webp))",
    re.IGNORECASE
)

def guess_mime(filename: str) -> str:
    mt, _ = mimetypes.guess_type(filename)
    return mt or "application/octet-stream"

def extract_image_filename_from_placeholder_text(p_text: str) -> Optional[str]:
    m = IMG_PLACEHOLDER_RE.search(p_text or "")
    if not m:
        return None
    return m.group(1)

def try_extract_size_from_placeholder_text(p_text: str) -> Tuple[Optional[int], Optional[int]]:
    # 예: "가로 2880pixel, 세로 1718pixel"
    w = None
    h = None
    m1 = re.search(r"가로\s*(\d+)\s*pixel", p_text)
    m2 = re.search(r"세로\s*(\d+)\s*pixel", p_text)
    if m1:
        w = int(m1.group(1))
    if m2:
        h = int(m2.group(1))
    return w, h

def extract_all_bindata_files(zf: zipfile.ZipFile, out_dir: Path) -> Dict[str, str]:
    """
    BinData 내 파일을 assets로 추출하고
    {파일명: 상대경로} 매핑을 반환.
    """
    assets_dir = out_dir / "assets"
    assets_dir.mkdir(parents=True, exist_ok=True)

    mapping: Dict[str, str] = {}  # filename -> "assets/filename"

    for name in zf.namelist():
        if not name.startswith("BinData/"):
            continue
        if name.endswith("/"):
            continue

        filename = Path(name).name
        data = zf.read(name)
        (assets_dir / filename).write_bytes(data)
        mapping[filename] = f"assets/{filename}"

    return mapping

def render_img_tag(src: str, alt: str = "", width: Optional[int] = None, height: Optional[int] = None) -> str:
    w_attr = f' width="{width}"' if isinstance(width, int) and width > 0 else ""
    h_attr = f' height="{height}"' if isinstance(height, int) and height > 0 else ""
    return f"<img class='hwpx-img' src='{esc(src)}' alt='{esc(alt)}'{w_attr}{h_attr} />"


# -------------------------
# 1) Scan cell-merge schema
# -------------------------

@dataclass
class MergeHint:
    attr_candidates: Dict[str, int]
    child_tag_candidates: Dict[str, int]
    sample_cells: List[Tuple[str, str]]  # (tc_path, short_xml)

def scan_merge_hints(section_root: etree._Element, max_samples: int = 8) -> MergeHint:
    attr_counts: Dict[str, int] = {}
    child_counts: Dict[str, int] = {}
    samples: List[Tuple[str, str]] = []

    # tc 후보: local-name()='tc'
    tcs = section_root.xpath(".//*[local-name()='tc']")
    for idx, tc in enumerate(tcs):
        # 1) 속성 후보
        for k, v in tc.attrib.items():
            key = k.split("}")[-1]  # namespace 제거
            # 병합 관련으로 흔히 등장하는 키워드
            if any(x in key.lower() for x in ["span", "merge", "grid", "row", "col", "vmerge", "hmerge"]):
                attr_counts[key] = attr_counts.get(key, 0) + 1

        # 2) 하위 태그 후보 (tc 안에서 span/merge 관련 태그)
        for ch in tc.iterdescendants():
            ln = local_name(ch).lower()
            if any(x in ln for x in ["span", "merge", "grid", "row", "col"]):
                child_counts[ln] = child_counts.get(ln, 0) + 1

        # 샘플 수집: 병합 관련 흔적이 있거나, 그냥 초반 몇 개
        if len(samples) < max_samples:
            short = etree.tostring(tc, encoding="unicode")
            short = short[:800] + ("..." if len(short) > 800 else "")
            samples.append((f"tc[{idx}]", short))

    return MergeHint(attr_counts, child_counts, samples)


# -------------------------
# 2) Styles -> CSS (minimal)
# -------------------------

@dataclass
class StyleMaps:
    char_styles: Dict[str, Dict[str, str]]   # id -> css props
    para_styles: Dict[str, Dict[str, str]]   # id -> css props

def parse_styles(zf: zipfile.ZipFile) -> StyleMaps:
    char_styles: Dict[str, Dict[str, str]] = {}
    para_styles: Dict[str, Dict[str, str]] = {}

    style_paths = [p for p in zf.namelist() if p.startswith("Styles/") and p.endswith(".xml")]
    for sp in style_paths:
        root = read_xml_from_zip(zf, sp)
        if root is None:
            continue

        for st in root.xpath(".//*[local-name()='style' or local-name()='Style']"):
            sid = st.get("id") or st.get("ID") or st.get("styleId") or st.get("name")
            if not sid:
                continue

            if st.xpath(".//*[local-name()='charPr' or local-name()='CharPr']"):
                props = {}
                if st.xpath(".//*[local-name()='bold' or local-name()='b']"):
                    props["font-weight"] = "700"
                if st.xpath(".//*[local-name()='italic' or local-name()='i']"):
                    props["font-style"] = "italic"
                char_styles[str(sid)] = props

            if st.xpath(".//*[local-name()='paraPr' or local-name()='ParaPr']"):
                props = {}
                align_node = st.xpath(".//*[contains(translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'align')]")
                if align_node:
                    v = align_node[0].get("val") or align_node[0].get("value")
                    if v in ["left", "center", "right", "justify"]:
                        props["text-align"] = v
                para_styles[str(sid)] = props

    return StyleMaps(char_styles=char_styles, para_styles=para_styles)


# -------------------------
# 3) Document-order renderer
# -------------------------

def get_merge_span_from_tc(tc: etree._Element) -> Tuple[int, int]:
    # 1) attribute 기반 후보
    attr_row = tc.get("rowSpan") or tc.get("rowspan") or tc.get("vSpan") or tc.get("vspan")
    attr_col = tc.get("colSpan") or tc.get("colspan") or tc.get("hSpan") or tc.get("hspan")

    def to_int(x: Optional[str]) -> Optional[int]:
        if not x:
            return None
        if x.isdigit():
            return int(x)
        return None

    rs = to_int(attr_row)
    cs = to_int(attr_col)

    # 2) child tag 기반 후보
    if rs is None or cs is None:
        span_nodes = tc.xpath(".//*[contains(translate(local-name(), 'ABCDEFGHIJKLMNOPQRSTUVWXYZ','abcdefghijklmnopqrstuvwxyz'), 'span')]")
        for sn in span_nodes:
            r = sn.get("row") or sn.get("rowSpan") or sn.get("r")
            c = sn.get("col") or sn.get("colSpan") or sn.get("c")
            r2 = to_int(r)
            c2 = to_int(c)
            if rs is None and r2:
                rs = r2
            if cs is None and c2:
                cs = c2

    return (rs or 1, cs or 1)

def render_tbl(tbl: etree._Element) -> str:
    rows = tbl.xpath(".//*[local-name()='tr']")
    trs = []
    for tr in rows:
        tcs = tr.xpath("./*[local-name()='tc']")
        tds = []
        for tc in tcs:
            rs, cs = get_merge_span_from_tc(tc)
            attrs = []
            if rs > 1:
                attrs.append(f'rowspan="{rs}"')
            if cs > 1:
                attrs.append(f'colspan="{cs}"')

            txt = text_content(tc)
            attr_str = (" " + " ".join(attrs)) if attrs else ""
            tds.append(f"<td{attr_str}>{esc(txt)}</td>")
        trs.append("<tr>" + "".join(tds) + "</tr>")
    return "<table class='hwpx-table'>" + "".join(trs) + "</table>"

def render_p(p: etree._Element, bindata_map: Dict[str, str]) -> str:
    txt = text_content(p)
    if not txt:
        return ""

    # 이미지 placeholder 감지 → <img>로 치환
    img_name = extract_image_filename_from_placeholder_text(txt)
    if img_name:
        src = bindata_map.get(img_name)
        if src:
            w, h = try_extract_size_from_placeholder_text(txt)
            return render_img_tag(src=src, alt=img_name, width=w, height=h)
        return f"<p class='hwpx-p muted'>(이미지 파일을 BinData에서 찾지 못함: {esc(img_name)})</p>"

    return f"<p class='hwpx-p'>{esc(txt)}</p>"

def render_section0_in_doc_order(section_root: etree._Element, bindata_map: Dict[str, str]) -> str:
    blocks: List[str] = []

    candidates = section_root.xpath(".//*[local-name()='tbl' or local-name()='p']")
    for el in candidates:
        ln = local_name(el)
        if ln == "p":
            if el.xpath("ancestor::*[local-name()='tbl']"):
                continue
            html_p = render_p(el, bindata_map=bindata_map)
            if html_p:
                blocks.append(html_p)
        elif ln == "tbl":
            blocks.append(render_tbl(el))

    return "\n".join(blocks) if blocks else "<p class='muted'>(p/tbl을 찾지 못했습니다)</p>"


# -------------------------
# Main: hwpx -> HTML
# -------------------------

def build_html(hwpx_path: str, output_html: str = "index.html") -> None:
    hwpx_path = str(hwpx_path)
    out_dir = Path(output_html).resolve().parent  # index.html이 생길 폴더

    with zipfile.ZipFile(hwpx_path, "r") as zf:
        section0 = read_xml_from_zip(zf, "Contents/section0.xml")
        if section0 is None:
            raise FileNotFoundError("Contents/section0.xml 을 찾거나 파싱하지 못했습니다.")

        # ✅ BinData 추출(assets/)
        bindata_map = extract_all_bindata_files(zf, out_dir=out_dir)

        # 1) 병합 힌트 스캔
        hint = scan_merge_hints(section0)

        # 2) 스타일 맵(최소) 파싱
        style_maps = parse_styles(zf)

        # 3) 문서 흐름 렌더링 (이미지 placeholder 치환 포함)
        rendered = render_section0_in_doc_order(section0, bindata_map=bindata_map)

        # 4) pretty xml (디버그용)
        pretty = etree.tostring(section0, pretty_print=True, encoding="unicode")

    extracted_style_debug = []
    extracted_style_debug.append("<h3>Extracted style maps (debug)</h3>")
    extracted_style_debug.append("<pre class='xml'>")
    extracted_style_debug.append(esc(f"char_styles keys: {list(style_maps.char_styles.keys())}\n"))
    extracted_style_debug.append(esc(f"para_styles keys: {list(style_maps.para_styles.keys())}\n"))
    extracted_style_debug.append("</pre>")

    merge_report = []
    merge_report.append("<h3>Cell-merge schema hints</h3>")
    merge_report.append("<pre class='xml'>")
    merge_report.append(esc("ATTR candidates (count):\n"))
    for k, v in sorted(hint.attr_candidates.items(), key=lambda x: -x[1]):
        merge_report.append(esc(f"- {k}: {v}\n"))
    merge_report.append(esc("\nCHILD TAG candidates (count):\n"))
    for k, v in sorted(hint.child_tag_candidates.items(), key=lambda x: -x[1]):
        merge_report.append(esc(f"- {k}: {v}\n"))
    merge_report.append(esc("\nSAMPLES (first tc snippets):\n"))
    for path, snip in hint.sample_cells:
        merge_report.append(esc(f"\n[{path}]\n{snip}\n"))
    merge_report.append("</pre>")

    html_doc = f"""
<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>HWPX section0 renderer</title>
  <style>
    body {{ background:#111; color:#eee; font-family:Consolas, monospace; padding:20px; }}
    h1,h2,h3 {{ color:#ffd479; }}
    .panel {{ background:#151515; border:1px solid #333; border-radius:10px; padding:16px; margin:12px 0; }}
    .muted {{ color:#888; }}
    .xml {{ background:#1e1e1e; padding:16px; overflow:auto; border-radius:8px; white-space:pre-wrap; }}
    .render {{ background:#151515; padding:16px; border-radius:8px; }}
    .hwpx-p {{ line-height:1.7; margin:6px 0; white-space:pre-wrap; }}

    .hwpx-img {{
      display: block;
      max-width: 100%;
      height: auto;
      margin: 10px 0 18px 0;
      border: 1px solid #333;
      border-radius: 8px;
      background: #0e0e0e;
    }}

    table.hwpx-table {{
      border-collapse: collapse;
      width: 100%;
      margin: 10px 0 18px 0;
      background:#101010;
    }}
    table.hwpx-table td {{
      border: 1px solid #444;
      padding: 8px;
      vertical-align: top;
      white-space: pre-wrap;
      min-width: 40px;
    }}
  </style>
</head>
<body>
  <h1>Contents/section0.xml Render (doc order)</h1>

  <div class="panel">
    <h2>Rendered</h2>
    <div class="render">{rendered}</div>
  </div>

  <div class="panel">
    {''.join(merge_report)}
  </div>

  <div class="panel">
    {''.join(extracted_style_debug)}
  </div>

  <div class="panel">
    <h2>section0.xml pretty</h2>
    <pre class="xml">{esc(pretty)}</pre>
  </div>
</body>
</html>
"""
    bindata = [n for n in zf.namelist() if n.startswith("BinData/") and not n.endswith("/")]
    print("BinData count:", len(bindata))
    print("\n".join(bindata[:80]))
    Path(output_html).write_text(html_doc, encoding="utf-8")


if __name__ == "__main__":
    # 예: template.hwpx를 같은 폴더에 두고 실행
    build_html("template.hwpx", "index.html")
