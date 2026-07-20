# -*- coding: utf-8 -*-
"""
招商银行交易流水PDF提取工具
从招商银行交易流水PDF中提取交易记录，导出为Excel文件。

使用 pdf_cmap_decoder 模块自动逆向解析PDF字体编码，
无需手动维护字符映射表，具有通用性。

依赖: pip install pymupdf fonttools pandas openpyxl
"""

import re
import sys
import os
import pandas as pd
from pdf_cmap_decoder import extract_text


def parse_transactions(pages_lines):
    """
    解析交易记录。

    参数:
        pages_lines: [[line_text, ...], ...] 每页的行列表

    每条记录的字段各占一行:
    日期 → CNY → 交易金额 → 余额 → 交易摘要 → 对手信息(可能多行)
    """
    transactions = []

    date_pattern = re.compile(r'^\d{4}-\d{2}-\d{2}$')
    amount_pattern = re.compile(r'^-?[\d,]+\.\d{2}$')
    page_num_pattern = re.compile(r'^\d+/\d+$')

    # 需要跳过的页头/页脚关键词
    skip_keywords = {
        'Transaction Statement', 'Name', 'Account Type',
        'Account No', 'Sub Branch', 'Verification Code',
        'Date', 'Currency', 'Transaction', 'Amount',
        'Balance', 'Counter Party', 'Transaction Type',
    }

    # 标题区域关键词(正确的中文)
    header_keywords = [
        '记账日期', '货币', '交易金额', '联机余额',
        '交易摘要', '对手信息', '招商银行交易流水',
        '户  名', '账户类型', '申请时间', '账号',
        '开 户 行', '验 证 码',
    ]

    current_record = None
    # 状态机: 0=等待日期, 1=等待币种, 2=等待金额, 3=等待余额, 4=等待摘要, 5=读取对手信息
    state = 0
    temp_date = None
    temp_amount = None
    temp_balance = None

    for page_lines in pages_lines:
        for line in page_lines:
            line = line.strip()
            if not line:
                continue

            # 跳过页码
            if page_num_pattern.match(line):
                continue

            # 跳过英文页头
            if line in skip_keywords:
                continue

            # 跳过中文标题区域
            if any(kw in line for kw in header_keywords):
                continue

            # 跳过日期范围行
            if '--' in line and re.search(r'\d{4}-\d{2}-\d{2}', line):
                continue

            if state == 0:
                if date_pattern.match(line):
                    temp_date = line
                    state = 1
                elif current_record:
                    # 上一条记录的对手信息续行
                    if current_record['对手信息']:
                        current_record['对手信息'] += ' ' + line
                    else:
                        current_record['对手信息'] = line

            elif state == 1:
                if line == 'CNY':
                    state = 2
                else:
                    state = 0

            elif state == 2:
                if amount_pattern.match(line):
                    temp_amount = float(line.replace(',', ''))
                    state = 3
                else:
                    state = 0

            elif state == 3:
                if amount_pattern.match(line):
                    temp_balance = float(line.replace(',', ''))
                    state = 4
                else:
                    state = 0

            elif state == 4:
                if date_pattern.match(line):
                    # 无摘要直接到了下一条日期
                    if current_record:
                        transactions.append(current_record)
                    current_record = {
                        '记账日期': temp_date,
                        '币种': 'CNY',
                        '交易金额': temp_amount,
                        '联机余额': temp_balance,
                        '交易摘要': '',
                        '对手信息': '',
                    }
                    temp_date = line
                    state = 1
                else:
                    if current_record:
                        transactions.append(current_record)
                    current_record = {
                        '记账日期': temp_date,
                        '币种': 'CNY',
                        '交易金额': temp_amount,
                        '联机余额': temp_balance,
                        '交易摘要': line,
                        '对手信息': '',
                    }
                    state = 5

            elif state == 5:
                if date_pattern.match(line):
                    temp_date = line
                    state = 1
                else:
                    if line in skip_keywords:
                        continue
                    if current_record['对手信息']:
                        current_record['对手信息'] += ' ' + line
                    else:
                        current_record['对手信息'] = line

    if current_record:
        transactions.append(current_record)

    return transactions


def _clean_illegal_chars(text):
    """移除Excel不允许的控制字符"""
    if not text:
        return text
    # 保留 \t(09) \n(0a) \r(0d)，移除其他控制字符
    return re.sub(r'[\x00-\x08\x0b\x0c\x0e-\x1f]', '', text)


def export_to_excel(transactions, output_path):
    """将交易记录导出为Excel"""
    # 清理非法字符
    for t in transactions:
        for key in t:
            if isinstance(t[key], str):
                t[key] = _clean_illegal_chars(t[key])

    df = pd.DataFrame(transactions, columns=[
        '记账日期', '币种', '交易金额', '联机余额', '交易摘要', '对手信息'
    ])

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='交易流水')

        worksheet = writer.sheets['交易流水']
        column_widths = {
            'A': 14, 'B': 8, 'C': 14,
            'D': 14, 'E': 22, 'F': 55,
        }
        for col, width in column_widths.items():
            worksheet.column_dimensions[col].width = width

    return len(df)


def main():
    pdf_path = r"d:\Hui-87F166\华为家庭存储\个人文档\金融和消费\招商银行\招商银行交易流水(200120～240916).pdf"

    output_dir = os.path.dirname(pdf_path)
    output_path = os.path.join(
        output_dir, "招商银行交易流水(200120～240916).xlsx"
    )

    if not os.path.exists(pdf_path):
        local_pdf = os.path.join(
            os.getcwd(), "招商银行交易流水(200120～240916).pdf"
        )
        if os.path.exists(local_pdf):
            pdf_path = local_pdf
            output_path = os.path.join(
                os.getcwd(), "招商银行交易流水(200120～240916).xlsx"
            )
        else:
            print(f"错误: 找不到PDF文件: {pdf_path}")
            sys.exit(1)

    print(f"正在读取PDF: {pdf_path}")

    # 1. 使用CMap逆向解析模块提取文本(自动修正字体编码)
    pages_lines = extract_text(pdf_path)
    total_lines = sum(len(p) for p in pages_lines)
    print(f"共提取 {len(pages_lines)} 页, {total_lines} 行文本")

    # 2. 解析交易记录
    transactions = parse_transactions(pages_lines)
    print(f"共解析 {len(transactions)} 条交易记录")

    if not transactions:
        print("警告: 未解析到任何交易记录，请检查PDF格式。")
        sys.exit(1)

    # 3. 导出Excel
    count = export_to_excel(transactions, output_path)
    print(f"成功导出 {count} 条记录到: {output_path}")

    # 打印前10条记录预览
    print(f"\n前10条记录预览:")
    print("-" * 110)
    print(f"{'日期':<12} {'金额':>12} {'余额':>12} {'摘要':<18} {'对手信息'}")
    print("-" * 110)
    for t in transactions[:10]:
        cp = t['对手信息'][:35] if t['对手信息'] else ''
        print(
            f"{t['记账日期']:<12} {t['交易金额']:>12,.2f} "
            f"{t['联机余额']:>12,.2f} {t['交易摘要']:<18} {cp}"
        )
    print("-" * 110)

    # 打印最后5条记录
    print(f"\n最后5条记录:")
    print("-" * 110)
    for t in transactions[-5:]:
        cp = t['对手信息'][:35] if t['对手信息'] else ''
        print(
            f"{t['记账日期']:<12} {t['交易金额']:>12,.2f} "
            f"{t['联机余额']:>12,.2f} {t['交易摘要']:<18} {cp}"
        )
    print("-" * 110)


if __name__ == '__main__':
    main()
