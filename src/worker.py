# -*- coding: utf-8 -*-
"""
招商银行交易流水提取 - 控制台工作进程
由 C# GUI 调用，通过 stdout 输出进度，exit code 表示结果。

用法: worker.exe <pdf_path> <output_xlsx_path>
"""

import os
import re
import sys
import io

# 强制 stdout/stderr 使用 UTF-8（Windows 控制台默认 GBK，C# 端按 UTF-8 读取）
if hasattr(sys.stdout, 'buffer'):
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
if hasattr(sys.stderr, 'buffer'):
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

from openpyxl import Workbook
from pdf_cmap_decoder import extract_text


def parse_transactions(pages_lines):
    transactions = []
    date_pat = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    amount_pat = re.compile(r'^-?[\d,]+\.\d{2}$')
    page_num_pat = re.compile(r'^\d+/\d+$')

    skip_kw = {
        'Transaction Statement', 'Name', 'Account Type',
        'Account No', 'Sub Branch', 'Verification Code',
        'Date', 'Currency', 'Transaction', 'Amount',
        'Balance', 'Counter Party', 'Transaction Type',
    }
    header_kw = [
        '记账日期', '货币', '交易金额', '联机余额',
        '交易摘要', '对手信息', '招商银行交易流水',
        '户  名', '账户类型', '申请时间', '账号',
        '开 户 行', '验 证 码',
    ]

    rec = None
    state = 0
    t_date = t_amt = t_bal = None

    for lines in pages_lines:
        for line in lines:
            line = line.strip()
            if not line:
                continue
            if page_num_pat.match(line):
                continue
            if line in skip_kw:
                continue
            if any(k in line for k in header_kw):
                continue
            if '--' in line and re.search(r'\d{4}-\d{2}-\d{2}', line):
                continue

            if state == 0:
                if date_pat.match(line):
                    t_date = line; state = 1
                elif rec:
                    rec['对手信息'] = (rec['对手信息'] + ' ' + line).strip()
            elif state == 1:
                state = 2 if line == 'CNY' else 0
            elif state == 2:
                if amount_pat.match(line):
                    t_amt = float(line.replace(',', '')); state = 3
                else:
                    state = 0
            elif state == 3:
                if amount_pat.match(line):
                    t_bal = float(line.replace(',', '')); state = 4
                else:
                    state = 0
            elif state == 4:
                if date_pat.match(line):
                    if rec: transactions.append(rec)
                    rec = {'记账日期': t_date, '币种': 'CNY',
                           '交易金额': t_amt, '联机余额': t_bal,
                           '交易摘要': '', '对手信息': ''}
                    t_date = line; state = 1
                else:
                    if rec: transactions.append(rec)
                    rec = {'记账日期': t_date, '币种': 'CNY',
                           '交易金额': t_amt, '联机余额': t_bal,
                           '交易摘要': line, '对手信息': ''}
                    state = 5
            elif state == 5:
                if date_pat.match(line):
                    t_date = line; state = 1
                elif line not in skip_kw:
                    rec['对手信息'] = (rec['对手信息'] + ' ' + line).strip()

    if rec:
        transactions.append(rec)
    return transactions


def export_to_excel(transactions, output_path):
    for t in transactions:
        for k in t:
            if isinstance(t[k], str):
                t[k] = re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', t[k])

    cols = ['记账日期', '币种', '交易金额', '联机余额', '交易摘要', '对手信息']
    wb = Workbook()
    ws = wb.active
    ws.title = '交易流水'
    ws.append(cols)
    for t in transactions:
        ws.append([t[c] for c in cols])
    for col, wd in {'A': 14, 'B': 8, 'C': 14, 'D': 14, 'E': 22, 'F': 55}.items():
        ws.column_dimensions[col].width = wd
    wb.save(output_path)
    return len(transactions)


def log(msg):
    """输出一行日志到 stdout（供 GUI 读取）"""
    print(msg, flush=True)


def main():
    if len(sys.argv) < 3:
        log("[ERROR] 用法: worker.exe <pdf路径> <输出xlsx路径>")
        sys.exit(1)

    pdf_path = sys.argv[1]
    output_path = sys.argv[2]

    if not os.path.isfile(pdf_path):
        log(f"[ERROR] PDF文件不存在: {pdf_path}")
        sys.exit(1)

    try:
        log(f"读取 PDF: {os.path.basename(pdf_path)}")
        log("正在解析 PDF 字体编码…")

        pages = extract_text(pdf_path)
        n = sum(len(p) for p in pages)
        log(f"提取 {len(pages)} 页, {n} 行文本")

        log("正在解析交易记录…")
        txns = parse_transactions(pages)
        log(f"解析 {len(txns)} 条交易记录")

        if not txns:
            log("[ERROR] 未解析到交易记录，请确认PDF为招商银行交易流水格式。")
            sys.exit(1)

        log("正在导出 Excel…")
        cnt = export_to_excel(txns, output_path)
        log(f"导出 {cnt} 条记录 → {output_path}")
        log(f"[DONE] {output_path}")
        sys.exit(0)

    except Exception as exc:
        log(f"[ERROR] {exc}")
        sys.exit(1)


if __name__ == '__main__':
    main()
