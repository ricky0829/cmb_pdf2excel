# -*- coding: utf-8 -*-
"""
pdf_cmap_decoder.py - PDF字体CMap逆向解析模块

解决银行PDF等使用子集化字体+打乱编码导致文本提取乱码的问题。

原理:
  PDF使用Identity-H编码时, 内容流中的字符代码(CID)通过CIDToGIDMap映射到
  嵌入字体的字形(GID)。当ToUnicode CMap不完整时, PyMuPDF等库会将CID直接
  当作Unicode输出, 导致乱码。

  本模块通过以下步骤还原正确文本:
  1. 提取PDF中的CIDToGIDMap和ToUnicode CMap
  2. 提取嵌入字体的字形轮廓
  3. 与参考字体(如SimSun)的字形轮廓进行特征匹配
  4. 建立完整的 CID → Unicode 映射
  5. 解析PDF内容流, 用正确映射解码文本

依赖: PyMuPDF(fitz), fonttools
"""

import re
import struct
import logging
from io import BytesIO
from pathlib import Path

import fitz
from fontTools.ttLib import TTFont
from fontTools.pens.recordingPen import RecordingPen

logger = logging.getLogger(__name__)


# ============================================================
#  第一部分: PDF字体结构提取
# ============================================================

def parse_tounicode_cmap(stream_bytes):
    """解析ToUnicode CMap流, 返回 {CID: Unicode} 映射"""
    text = stream_bytes.decode('latin-1')
    mapping = {}

    # bfchar: <CID> <Unicode>
    for m in re.finditer(
        r'beginbfchar\s*(.*?)\s*endbfchar', text, re.DOTALL
    ):
        for pair in re.finditer(r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>', m.group(1)):
            mapping[int(pair.group(1), 16)] = int(pair.group(2), 16)

    # bfrange: <start> <end> <unicodeStart>
    for m in re.finditer(
        r'beginbfrange\s*(.*?)\s*endbfrange', text, re.DOTALL
    ):
        for triple in re.finditer(
            r'<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>\s*<([0-9A-Fa-f]+)>',
            m.group(1),
        ):
            s, e, u = (int(triple.group(i), 16) for i in (1, 2, 3))
            for i in range(e - s + 1):
                mapping[s + i] = u + i

    return mapping


def _get_cid_to_gid_map(doc, xref):
    """提取CIDToGIDMap, 返回 {CID: GID}; Identity映射返回None"""
    obj = doc.xref_object(xref)
    m = re.search(r'/CIDToGIDMap\s+(\d+)\s+0\s+R', obj)
    if m:
        data = doc.xref_stream(int(m.group(1)))
        n = len(data) // 2
        return {
            cid: struct.unpack('>H', data[cid * 2: cid * 2 + 2])[0]
            for cid in range(n)
        }
    if '/CIDToGIDMap /Identity' in obj:
        return None  # CID == GID
    return None


def _find_fontfile_xref(doc, font_descriptor_xref):
    """从FontDescriptor中找到嵌入字体流的xref"""
    fd_obj = doc.xref_object(font_descriptor_xref)
    for key in ('/FontFile2', '/FontFile', '/FontFile3'):
        m = re.search(rf'{key}\s+(\d+)\s+0\s+R', fd_obj)
        if m:
            return int(m.group(1))
    return None


def collect_font_info(doc, page):
    """
    收集页面所有字体的映射信息。

    返回 {font_name: {
        'xref', 'base_font', 'encoding',
        'tounicode': {CID: Unicode},
        'cid_to_gid': {CID: GID} | None,
        'descendant_xref': int,
        'fontfile_xref': int | None,
    }}
    """
    fonts = {}
    page_obj = doc.xref_object(page.xref)

    # 定位 /Font 资源字典: Page → Resources → Font
    font_dict_xref = None

    # 方式1: Resources是间接引用
    res_m = re.search(r'/Resources\s+(\d+)\s+0\s+R', page_obj)
    if res_m:
        res_obj = doc.xref_object(int(res_m.group(1)))
        fm = re.search(r'/Font\s+(\d+)\s+0\s+R', res_obj)
        if fm:
            font_dict_xref = int(fm.group(1))
        else:
            # Font可能是内联字典
            fm2 = re.search(r'/Font\s*<<(.+?)>>', res_obj, re.DOTALL)
            if fm2:
                font_dict_str = fm2.group(1)
                for fe in re.finditer(r'/(\w+)\s+(\d+)\s+0\s+R', font_dict_str):
                    fonts[fe.group(1)] = _parse_font_entry(
                        doc, int(fe.group(2))
                    )
                return fonts

    # 方式2: Font直接在页面对象中
    if font_dict_xref is None:
        fm = re.search(r'/Font\s+(\d+)\s+0\s+R', page_obj)
        if fm:
            font_dict_xref = int(fm.group(1))

    if font_dict_xref is None:
        return fonts

    font_dict_str = doc.xref_object(font_dict_xref)

    for fm2 in re.finditer(r'/(\w+)\s+(\d+)\s+0\s+R', font_dict_str):
        font_name = fm2.group(1)
        font_xref = int(fm2.group(2))
        fonts[font_name] = _parse_font_entry(doc, font_xref)

    return fonts


def _parse_font_entry(doc, font_xref):
    """解析单个字体对象的完整映射信息"""
    font_obj = doc.xref_object(font_xref)

    base_m = re.search(r'/BaseFont\s*/(\S+)', font_obj)
    enc_m = re.search(r'/Encoding\s*/(\S+)', font_obj)

    info = {
        'xref': font_xref,
        'base_font': base_m.group(1) if base_m else '',
        'encoding': enc_m.group(1) if enc_m else '',
        'tounicode': {},
        'cid_to_gid': None,
        'descendant_xref': None,
        'fontfile_xref': None,
    }

    # ToUnicode
    tu_m = re.search(r'/ToUnicode\s+(\d+)\s+0\s+R', font_obj)
    if tu_m:
        try:
            info['tounicode'] = parse_tounicode_cmap(
                doc.xref_stream(int(tu_m.group(1)))
            )
        except Exception:
            pass

    # DescendantFonts → CIDToGIDMap + FontDescriptor
    desc_m = re.search(r'/DescendantFonts\s*\[\s*(\d+)\s+0\s+R', font_obj)
    if desc_m:
        desc_xref = int(desc_m.group(1))
        info['descendant_xref'] = desc_xref
        info['cid_to_gid'] = _get_cid_to_gid_map(doc, desc_xref)

        desc_obj = doc.xref_object(desc_xref)
        fd_m = re.search(r'/FontDescriptor\s+(\d+)\s+0\s+R', desc_obj)
        if fd_m:
            info['fontfile_xref'] = _find_fontfile_xref(
                doc, int(fd_m.group(1))
            )

    return info


# ============================================================
#  第二部分: 字形轮廓匹配
# ============================================================

def _glyph_signature(glyf_table, glyph_name):
    """获取glyph的轮廓特征 (numberOfContours, bbox)"""
    g = glyf_table[glyph_name]
    if g.numberOfContours == 0:
        return None
    if not hasattr(g, 'xMin'):
        return None
    return (g.numberOfContours, g.xMin, g.yMin, g.xMax, g.yMax)


def _build_reference_index(ref_font):
    """为参考字体建立 (contours, bbox) → [Unicode] 索引"""
    glyf = ref_font['glyf']
    cmap = ref_font['cmap'].getBestCmap()
    index = {}

    for uni, gname in cmap.items():
        sig = _glyph_signature(glyf, gname)
        if sig:
            index.setdefault(sig, []).append(uni)

    return index


def _compare_glyph_outlines(glyf1, name1, glyf2, name2):
    """比较两个glyph的轮廓坐标数据是否一致(支持简单和复合字形)"""
    g1 = glyf1[name1]
    g2 = glyf2[name2]

    # 复合glyph: 先尝试展开为轮廓再比较
    if g1.isComposite() or g2.isComposite():
        return _compare_via_pen(glyf1, name1, glyf2, name2)

    # 简单glyph: 直接比较坐标数据
    if g1.numberOfContours != g2.numberOfContours:
        return False
    if g1.numberOfContours <= 0:
        return g1.numberOfContours == g2.numberOfContours

    try:
        c1, f1, e1 = g1.getCoordinates(glyf1)
        c2, f2, e2 = g2.getCoordinates(glyf2)
    except Exception:
        return _compare_via_pen(glyf1, name1, glyf2, name2)

    if len(c1) != len(c2) or list(e1) != list(e2):
        return False

    for (x1, y1), (x2, y2) in zip(c1, c2):
        if x1 != x2 or y1 != y2:
            return False

    return True


def _compare_via_pen(glyf1, name1, glyf2, name2):
    """使用RecordingPen比较glyph轮廓(用于复合字形)"""
    pen1, pen2 = RecordingPen(), RecordingPen()
    try:
        glyf1[name1].draw(pen1)
        glyf2[name2].draw(pen2)
    except Exception:
        return False

    if len(pen1.value) != len(pen2.value):
        return False

    for (op1, args1), (op2, args2) in zip(pen1.value, pen2.value):
        if op1 != op2:
            return False
        if op1 in ('endPath', 'closePath'):
            continue
        for a1, a2 in zip(args1, args2):
            if isinstance(a1, tuple) and isinstance(a2, tuple):
                if len(a1) != len(a2):
                    return False
                for v1, v2 in zip(a1, a2):
                    if abs(v1 - v2) > 1:
                        return False
            elif a1 != a2:
                return False

    return True


def build_glyph_mapping(embedded_font_bytes, ref_font, existing_tounicode=None):
    """
    通过字形轮廓匹配建立 GID → Unicode 映射。

    参数:
        embedded_font_bytes: 嵌入字体的二进制数据
        ref_font: 参考字体 TTFont 对象
        existing_tounicode: 已有的ToUnicode映射(用于验证), 可选

    返回: {GID: Unicode}
    """
    embedded = TTFont(BytesIO(embedded_font_bytes))
    emb_glyphs = embedded.getGlyphOrder()
    emb_glyf = embedded['glyf']
    num_glyphs = len(emb_glyphs)

    # 建立参考字体索引
    ref_index = _build_reference_index(ref_font)
    ref_glyf = ref_font['glyf']
    ref_cmap = ref_font['cmap'].getBestCmap()
    # ref_cmap: {Unicode: glyph_name}

    gid_to_unicode = {}
    matched = 0
    ambiguous = 0
    unmatched = 0

    for gid in range(num_glyphs):
        gname = emb_glyphs[gid]
        sig = _glyph_signature(emb_glyf, gname)
        if sig is None:
            continue  # 空glyph (.notdef, space等)

        candidates = ref_index.get(sig, [])

        if len(candidates) == 1:
            gid_to_unicode[gid] = candidates[0]
            matched += 1
        elif len(candidates) > 1:
            # 多候选: 精确比较轮廓
            found = False
            for uni in candidates:
                ref_gname = ref_cmap.get(uni)
                if ref_gname and _compare_glyph_outlines(
                    emb_glyf, gname, ref_glyf, ref_gname
                ):
                    gid_to_unicode[gid] = uni
                    matched += 1
                    found = True
                    break
            if not found:
                ambiguous += 1
                logger.debug(f"GID {gid}: {len(candidates)} candidates, no exact match")
        else:
            unmatched += 1
            logger.debug(f"GID {gid}: no match in reference font")

    logger.info(
        f"Glyph matching: {matched} matched, {ambiguous} ambiguous, "
        f"{unmatched} unmatched (total {num_glyphs})"
    )

    embedded.close()
    return gid_to_unicode


# ============================================================
#  第三部分: 内容流解析
# ============================================================

def _decode_hex_string(hex_str):
    """解码十六进制字符串为CID列表"""
    hex_str = hex_str.strip()
    if len(hex_str) % 4 != 0:
        hex_str = hex_str.ljust((len(hex_str) + 3) // 4 * 4, '0')
    return [int(hex_str[i:i + 4], 16) for i in range(0, len(hex_str), 4)]


def _decode_literal_string(s):
    """解码字面字符串为CID列表"""
    cids = []
    i = 0
    while i < len(s):
        if s[i] == '\\' and i + 1 < len(s):
            c = s[i + 1]
            if c == 'n':
                cids.append(10); i += 2
            elif c == 'r':
                cids.append(13); i += 2
            elif c == 't':
                cids.append(9); i += 2
            elif c == '(':
                cids.append(40); i += 2
            elif c == ')':
                cids.append(41); i += 2
            elif c == '\\':
                cids.append(92); i += 2
            elif c.isdigit():
                # 八进制
                octal = c
                i += 2
                while i < len(s) and s[i].isdigit() and len(octal) < 3:
                    octal += s[i]; i += 1
                cids.append(int(octal, 8))
            else:
                cids.append(ord(c)); i += 2
        else:
            cids.append(ord(s[i])); i += 1
    return cids


def _extract_string_arg(arg_str):
    """从操作符参数中提取字符串内容"""
    arg_str = arg_str.strip()
    if arg_str.startswith('<') and arg_str.endswith('>'):
        return 'hex', arg_str[1:-1]
    if arg_str.startswith('(') and arg_str.endswith(')'):
        return 'literal', arg_str[1:-1]
    return None, None


def parse_content_stream(doc, page, font_mappings):
    """
    解析PDF页面内容流, 使用正确的字体映射解码文本。

    返回: [line_text, ...] 按页面位置排列的文本行
    """
    lines = []
    current_chars = []

    # 文本状态
    in_text = False
    current_font = None
    font_size = 10.0
    tm = [1, 0, 0, 1, 0, 0]   # 文本矩阵
    tlm = [1, 0, 0, 1, 0, 0]  # 行矩阵
    leading = 0.0

    def flush_line():
        nonlocal current_chars
        if current_chars:
            text = ''.join(c[0] for c in current_chars).rstrip()
            if text:
                lines.append(text)
        current_chars = []

    def cid_to_char(cid, font_info):
        """将CID转换为正确的Unicode字符"""
        # 优先使用ToUnicode CMap
        if cid in font_info['tounicode']:
            return chr(font_info['tounicode'][cid])

        # 通过CIDToGIDMap → 字形匹配
        if font_info.get('glyph_map'):
            c2g = font_info['cid_to_gid']
            if c2g is not None:
                gid = c2g.get(cid, 0)
            else:
                gid = cid  # Identity
            uni = font_info['glyph_map'].get(gid)
            if uni is not None:
                return chr(uni)

        # 最终回退: CID作为Unicode
        return chr(cid) if cid > 0 else ''

    def process_show_string(arg_str):
        """处理Tj操作符的字符串"""
        nonlocal current_chars
        str_type, content = _extract_string_arg(arg_str)
        if str_type is None or current_font is None:
            return

        font_info = font_mappings.get(current_font)
        if font_info is None:
            return

        if str_type == 'hex':
            cids = _decode_hex_string(content)
        else:
            cids = _decode_literal_string(content)

        for cid in cids:
            ch = cid_to_char(cid, font_info)
            if ch:
                x = tm[4]
                y = tm[5]
                current_chars.append((ch, x, y))
                # 简单推进x坐标(精确值需要glyph宽度, 这里用近似)
                tm[4] += font_size * 0.5

    def process_tj_array(array_str):
        """处理TJ操作符的数组"""
        # 提取数组中的字符串元素
        for m in re.finditer(r'<([0-9A-Fa-f]+)>|\(([^)]*)\)', array_str):
            if m.group(1) is not None:
                process_show_string(f"<{m.group(1)}>")
            elif m.group(2) is not None:
                process_show_string(f"({m.group(2)})")

    # 逐行处理内容流
    for cx in page.get_contents():
        raw = doc.xref_stream(cx)
        text = raw.decode('latin-1')

        for line in text.split('\n'):
            line = line.strip()
            if not line:
                continue

            if line == 'BT':
                in_text = True
                tm = [1, 0, 0, 1, 0, 0]
                tlm = [1, 0, 0, 1, 0, 0]
                continue

            if line == 'ET':
                flush_line()
                in_text = False
                continue

            if not in_text:
                continue

            # Tf - 设置字体
            m = re.match(r'/(\w+)\s+([\d.]+)\s+Tf', line)
            if m:
                current_font = m.group(1)
                font_size = float(m.group(2))
                continue

            # Tm - 设置文本矩阵
            m = re.match(
                r'([\d.e+-]+)\s+([\d.e+-]+)\s+([\d.e+-]+)\s+'
                r'([\d.e+-]+)\s+([\d.e+-]+)\s+([\d.e+-]+)\s+Tm',
                line,
            )
            if m:
                flush_line()
                vals = [float(m.group(i)) for i in range(1, 7)]
                tm = list(vals)
                tlm = list(vals)
                continue

            # Td / TD - 移动文本位置
            m = re.match(r'([\d.e+-]+)\s+([\d.e+-]+)\s+(Td|TD)', line)
            if m:
                flush_line()
                tx, ty = float(m.group(1)), float(m.group(2))
                if m.group(3) == 'TD':
                    leading = -ty
                tlm[4] += tx * tlm[0] + ty * tlm[2]
                tlm[5] += tx * tlm[1] + ty * tlm[3]
                tm = list(tlm)
                continue

            # T* - 下一行
            if line == 'T*':
                flush_line()
                tlm[5] -= leading
                tm = list(tlm)
                continue

            # TL - 设置行距
            m = re.match(r'([\d.e+-]+)\s+TL', line)
            if m:
                leading = float(m.group(1))
                continue

            # Tj - 显示字符串
            m = re.match(r'(<[0-9A-Fa-f]+>|\([^)]*\))\s+Tj', line)
            if m:
                process_show_string(m.group(1))
                continue

            # TJ - 显示字符串数组
            m = re.match(r'\[(.+)\]\s+TJ', line)
            if m:
                process_tj_array(m.group(1))
                continue

            # ' - 下一行并显示字符串
            m = re.match(r'(<[0-9A-Fa-f]+>|\([^)]*\))\s+\'', line)
            if m:
                flush_line()
                tlm[5] -= leading
                tm = list(tlm)
                process_show_string(m.group(1))
                continue

    flush_line()
    return lines


# ============================================================
#  第四部分: 主入口
# ============================================================

def _find_reference_font(font_name):
    """根据PDF字体名自动查找系统参考字体"""
    name_lower = font_name.lower()

    # Windows字体目录
    win_fonts = Path(r'C:\Windows\Fonts')

    candidates = []
    if 'simsun' in name_lower or 'song' in name_lower:
        candidates = [
            win_fonts / 'simsun.ttc',
            win_fonts / 'SIMSUN.TTC',
        ]
    elif 'simhei' in name_lower or 'hei' in name_lower:
        candidates = [win_fonts / 'simhei.ttf', win_fonts / 'SIMHEI.TTF']
    elif 'simkai' in name_lower or 'kai' in name_lower:
        candidates = [win_fonts / 'simkai.ttf', win_fonts / 'SIMKAI.TTF']
    elif 'simfang' in name_lower or 'fang' in name_lower:
        candidates = [win_fonts / 'simfang.ttf', win_fonts / 'SIMFANG.TTF']
    elif 'msyh' in name_lower or 'yahei' in name_lower:
        candidates = [win_fonts / 'msyh.ttc', win_fonts / 'MSYH.TTC']

    for p in candidates:
        if p.exists():
            return str(p)

    # 默认尝试SimSun
    default = win_fonts / 'simsun.ttc'
    if default.exists():
        return str(default)

    return None


def extract_text(pdf_path, ref_font_path=None, verbose=False):
    """
    从PDF中提取文本, 自动修正字体编码乱码。

    参数:
        pdf_path: PDF文件路径
        ref_font_path: 参考字体路径(可选, 默认自动检测SimSun)
        verbose: 是否输出调试信息

    返回:
        [page_lines, ...] 其中 page_lines = [line_text, ...]
    """
    if verbose:
        logging.basicConfig(level=logging.DEBUG)
    else:
        logging.basicConfig(level=logging.INFO)

    doc = fitz.open(pdf_path)
    num_pages = len(doc)
    logger.info(f"PDF: {num_pages} pages")

    # ---- 收集所有页面的字体信息 ----
    all_fonts = {}  # font_xref → font_info (跨页面共享)
    page_font_names = []  # 每页的 {name: xref}

    for page in doc:
        fonts = collect_font_info(doc, page)
        name_map = {}
        for name, info in fonts.items():
            xref = info['xref']
            if xref not in all_fonts:
                all_fonts[xref] = info
            name_map[name] = xref
        page_font_names.append(name_map)

    logger.info(f"Found {len(all_fonts)} unique fonts")

    # ---- 建立字形匹配映射 ----
    if ref_font_path is None:
        # 自动检测参考字体
        for info in all_fonts.values():
            ref_font_path = _find_reference_font(info['base_font'])
            if ref_font_path:
                break

    ref_font = None
    if ref_font_path:
        try:
            ref_font = TTFont(ref_font_path, fontNumber=0)
            logger.info(f"Reference font: {ref_font_path}")
        except Exception as e:
            logger.warning(f"Cannot load reference font: {e}")

    for xref, info in all_fonts.items():
        info['glyph_map'] = {}  # GID → Unicode

        if ref_font and info.get('fontfile_xref'):
            try:
                font_bytes = doc.xref_stream(info['fontfile_xref'])
                gid_map = build_glyph_mapping(
                    font_bytes, ref_font, info['tounicode']
                )
                info['glyph_map'] = gid_map
                logger.info(
                    f"Font '{info['base_font']}': "
                    f"{len(info['tounicode'])} ToUnicode + "
                    f"{len(gid_map)} glyph-matched entries"
                )
            except Exception as e:
                logger.warning(f"Glyph matching failed for {info['base_font']}: {e}")

    # ---- 构建每页的字体映射(按字体名索引) ----
    page_mappings = []
    for name_map in page_font_names:
        mapping = {}
        for name, xref in name_map.items():
            mapping[name] = all_fonts[xref]
        page_mappings.append(mapping)

    # ---- 逐页解析内容流 ----
    all_pages = []
    for i, page in enumerate(doc):
        lines = parse_content_stream(doc, page, page_mappings[i])
        all_pages.append(lines)
        if verbose and i < 3:
            logger.debug(f"Page {i + 1}: {len(lines)} lines")

    if ref_font:
        ref_font.close()
    doc.close()

    logger.info("Extraction complete")
    return all_pages


# ============================================================
#  命令行入口
# ============================================================

if __name__ == '__main__':
    import sys

    if len(sys.argv) < 2:
        print("Usage: python pdf_cmap_decoder.py <pdf_path> [ref_font_path]")
        sys.exit(1)

    pdf = sys.argv[1]
    ref = sys.argv[2] if len(sys.argv) > 2 else None

    pages = extract_text(pdf, ref, verbose=True)
    for i, lines in enumerate(pages[:3]):
        print(f"\n{'='*60}")
        print(f"Page {i + 1} ({len(lines)} lines)")
        print('=' * 60)
        for line in lines[:30]:
            print(f"  {line}")
