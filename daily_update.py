#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
每日自动更新脚本：
  1. 从【现券市场交易情况汇总跟踪】提取缺失日度数据
  2. 更新【机构行为数据.xlsx】的8个Sheet
  3. 导出 data.json
  4. 提示部署

使用方法：
  python daily_update.py

回滚（撤销上次更新）：
  python daily_update.py --rollback 7
"""

import os
import sys
import io
import glob
import json
import re
import shutil
import smtplib
from email.mime.text import MIMEText

# 修复Windows终端GBK编码问题（防止print ✓等字符时崩溃）
# 使用安全print替代全局sys.stdout包装，避免与subprocess/atexit冲突
import builtins
_orig_print = builtins.print
def _safe_print(*args, **kwargs):
    try:
        _orig_print(*args, **kwargs)
    except UnicodeEncodeError:
        safe_args = [str(a).encode('gbk', 'replace').decode('gbk', 'replace') for a in args]
        _orig_print(*safe_args, **kwargs)
builtins.print = _safe_print


# === 邮件通知配置 ===
NOTIFY_EMAIL = '906881211@qq.com'
NOTIFY_AUTH = 'jfblwqnjdzgqbajh'
SMTP_SERVER = 'smtp.qq.com'
SMTP_PORT = 465


def send_notification(subject, body):
    """发送执行结果通知邮件（静默失败不影响主流程）"""
    try:
        msg = MIMEText(body, 'plain', 'utf-8')
        msg['Subject'] = subject
        msg['From'] = NOTIFY_EMAIL
        msg['To'] = NOTIFY_EMAIL
        with smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT) as s:
            s.login(NOTIFY_EMAIL, NOTIFY_AUTH)
            s.sendmail(NOTIFY_EMAIL, [NOTIFY_EMAIL], msg.as_string())
        print(f"  ✓ 通知邮件已发送")
    except Exception as e:
        print(f"  ⚠ 通知邮件发送失败: {e}")
import time
import datetime
import warnings
import subprocess
import win32com.client
import pythoncom

warnings.filterwarnings('ignore')

# === 配置 ===
TARGET_FILE = '机构行为数据.xlsx'
SOURCE_PATTERN = '现券市场交易情况汇总跟踪*.xlsx'

# 源Sheet名 → 目标Sheet名
SHEET_MAPPING = {
    '国债': '债元',
    '政金债': '金元',
    '地方政府债': '地元',
    '中票': '票',
    '短融超短融': '融',
    '企业债': '企',
    '其他': '其元',
}
# 信元 不在SHEET_MAPPING中，它=票+融+企（公式计算）

# 源文件列偏移：A=年份, B=月份, C=日期, D~BN=数据
# 目标文件：A=日期, B~BL=数据
# 目标列 = 源列 - 2
SOURCE_COL_OFFSET = 2


def find_source_file():
    """查找最新的源Excel文件"""
    matches = glob.glob(SOURCE_PATTERN)
    if not matches:
        print("⚠ 找不到源文件（现券市场交易情况汇总跟踪*.xlsx）")
        sys.exit(1)
    # 按修改时间排序，取最新
    matches.sort(key=os.path.getmtime, reverse=True)
    return os.path.abspath(matches[0])


def find_daily_start(ws):
    """在源Sheet的C列中查找'日期'标记行，返回数据起始行"""
    for r in range(3, 2000):
        val = ws.Cells(r, 3).Value
        if val is None:
            continue
        if isinstance(val, str) and '日期' in val:
            return r + 1
    return None


def get_date_value(ws, row):
    """获取源Sheet某行的日期值（C列）"""
    val = ws.Cells(row, 3).Value
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if hasattr(val, 'date'):
        return val.date()
    return None


def get_target_date(ws, row):
    """获取目标Sheet某行的日期（A列）"""
    val = ws.Cells(row, 1).Value
    if val is None:
        return None
    if isinstance(val, datetime.datetime):
        return val.date()
    if hasattr(val, 'date'):
        return val.date()
    return None


def get_last_col_in_row(ws, row):
    """获取指定行中最后一个有数据的列号"""
    return ws.Cells(row, ws.Columns.Count).End(-4159).Column


def update(target_file, source_file):
    """主更新逻辑"""
    tgt_path = os.path.abspath(target_file)
    src_path = source_file
    tmp_path = None

    print("=" * 60)
    print("  每日自动更新：现券市场 → 机构行为数据")
    print("=" * 60)
    print(f"源文件：{os.path.basename(src_path)}")
    print(f"目标文件：{os.path.basename(tgt_path)}")

    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        print("正在打开Excel文件...")
        src_wb = excel.Workbooks.Open(src_path)
        tgt_wb = excel.Workbooks.Open(tgt_path)
        excel.Calculation = -4135  # xlCalculationManual
        print("  文件已打开\n")

        max_inserted = 0

        for src_name, tgt_name in SHEET_MAPPING.items():
            print(f"【{src_name}】→【{tgt_name}】")

            src_sheets = [s.Name for s in src_wb.Sheets]
            tgt_sheets = [s.Name for s in tgt_wb.Sheets]

            if src_name not in src_sheets:
                print(f"  ⚠ 源Sheet不存在，跳过")
                continue
            if tgt_name not in tgt_sheets:
                print(f"  ⚠ 目标Sheet不存在，跳过")
                continue

            src_ws = src_wb.Sheets(src_name)
            tgt_ws = tgt_wb.Sheets(tgt_name)

            # 1. 找源Sheet日度数据起始行
            daily_start = find_daily_start(src_ws)
            if daily_start is None:
                print(f"  ⚠ 找不到日度数据起始行，跳过")
                continue

            # 2. 找源Sheet日度数据结束行
            src_last_row = daily_start
            while get_date_value(src_ws, src_last_row + 1) is not None:
                src_last_row += 1
            print(f"  源日度数据：行 {daily_start} ~ {src_last_row}")

            # 3. 获取目标Sheet最新日期（Row 3）
            tgt_latest = get_target_date(tgt_ws, 3)
            if tgt_latest is None:
                print(f"  ⚠ 目标Sheet第3行无日期，跳过")
                continue
            print(f"  目标最新日期：{tgt_latest}")

            # 4. 找出比目标更新的缺失行
            missing_rows = []
            for r in range(daily_start, src_last_row + 1):
                src_date = get_date_value(src_ws, r)
                if src_date is None:
                    break
                if src_date > tgt_latest:
                    missing_rows.append(r)
                else:
                    break

            n = len(missing_rows)
            if n == 0:
                print(f"  ✓ 无缺失数据\n")
                continue

            newest = get_date_value(src_ws, missing_rows[0])
            oldest = get_date_value(src_ws, missing_rows[-1])
            print(f"  需插入 {n} 行：{oldest} ~ {newest}")

            # 5. 列映射：源C~最后列 → 目标A~对应列
            src_header_row = daily_start - 1
            src_last_col = get_last_col_in_row(src_ws, src_header_row)
            total_cols = src_last_col - SOURCE_COL_OFFSET
            print(f"  列范围：源C~列{src_last_col} → 目标A~列{total_cols}")

            # 6. 在目标Row 3上方插入N行
            for i in range(n):
                tgt_ws.Rows(3).Insert()

            # 7. 复制数据（带列映射）
            for i, src_row in enumerate(missing_rows):
                tgt_row = 3 + i
                for j in range(total_cols):
                    src_col = 3 + j
                    tgt_col = 1 + j
                    tgt_ws.Cells(tgt_row, tgt_col).Value = src_ws.Cells(src_row, src_col).Value

            # 8. 复制格式
            fmt_row = 3 + n
            for new_row in range(3, 3 + n):
                tgt_ws.Rows(fmt_row).Copy()
                tgt_ws.Rows(new_row).PasteSpecial(-4122)  # xlPasteFormats
            tgt_wb.Application.CutCopyMode = False

            max_inserted = max(max_inserted, n)
            print(f"  ✓ 插入 {n} 行\n")

        # === 处理信元 = 票 + 融 + 企 ===
        print("【信元】= 票 + 融 + 企（公式）")
        tgt_sheets = [s.Name for s in tgt_wb.Sheets]
        if '信元' in tgt_sheets and all(s in tgt_sheets for s in ['票', '融', '企']):
            n = max_inserted
            if n > 0:
                xin_ws = tgt_wb.Sheets('信元')
                piao_ws = tgt_wb.Sheets('票')
                rong_ws = tgt_wb.Sheets('融')
                qi_ws = tgt_wb.Sheets('企')

                # 在信元Row 3上方插入N行
                for i in range(n):
                    xin_ws.Rows(3).Insert()

                # 获取数据列数
                last_col = piao_ws.Cells(2, piao_ws.Columns.Count).End(-4159).Column

                # 设置公式：信元 = 票 + 融 + 企
                for new_row in range(3, 3 + n):
                    for c in range(1, last_col + 1):
                        col_letter = get_col_letter(c)
                        if c == 1:
                            # A列=日期，从票Sheet取
                            xin_ws.Cells(new_row, c).Value = piao_ws.Cells(new_row, c).Value
                        else:
                            formula = f"=票!{col_letter}{new_row}+融!{col_letter}{new_row}+企!{col_letter}{new_row}"
                            xin_ws.Cells(new_row, c).Formula = formula

                # 复制格式
                fmt_row = 3 + n
                for new_row in range(3, 3 + n):
                    xin_ws.Rows(fmt_row).Copy()
                    xin_ws.Rows(new_row).PasteSpecial(-4122)
                tgt_wb.Application.CutCopyMode = False

                print(f"  ✓ 插入 {n} 行（公式）\n")
            else:
                print(f"  ✓ 无需更新\n")
        else:
            print(f"  ⚠ 缺少必要Sheet\n")

        # 保存（使用SaveAs替代Save，解决Save不持久化的问题）
        print("保存中...")
        import shutil
        tmp_path = tgt_path + '.tmp.xlsx'
        tgt_wb.SaveAs(tmp_path)
        print(f"\n{'=' * 60}")
        print(f"✓ 更新完成！共插入 {max_inserted} 行")
        print(f"{'=' * 60}")

        return max_inserted, tmp_path

    finally:
        try:
            src_wb.Close(SaveChanges=False)
        except:
            pass
        try:
            tgt_wb.Close(SaveChanges=False)
        except:
            pass
        try:
            excel.Quit()
        except:
            pass
        import gc
        del src_wb, tgt_wb, excel
        gc.collect()
        pythoncom.CoUninitialize()
        time.sleep(3)  # 等待Excel进程退出

        # 强制杀死残留Excel进程（不需要管理员权限，因为是同用户启动的）
        try:
            subprocess.run(['taskkill', '/IM', 'EXCEL.EXE', '/F'],
                           capture_output=True, timeout=10)
        except Exception:
            pass
        time.sleep(3)  # 等待进程完全退出、文件锁释放

        # 用临时文件替换原文件（带重试）
        if tmp_path and os.path.exists(tmp_path):
            for attempt in range(5):
                try:
                    if os.path.exists(tgt_path):
                        os.remove(tgt_path)
                    os.rename(tmp_path, tgt_path)
                    print("✓ 文件已保存")
                    break
                except PermissionError:
                    print(f"  ⚠ 文件仍被占用，{5 - attempt}s 后重试...")
                    time.sleep(5)
            else:
                print("  ✗ 无法替换原文件，临时文件保留为:", os.path.basename(tmp_path))


def get_col_letter(col_num):
    """列号转字母（1=A, 2=B, ..., 26=Z, 27=AA）"""
    result = ''
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def rollback(n):
    """回滚：删除上次插入的N行"""
    print("=" * 60)
    print("  回滚模式")
    print("=" * 60)

    tgt_path = os.path.abspath(TARGET_FILE)

    pythoncom.CoInitialize()
    excel = win32com.client.Dispatch("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False

    try:
        tgt_wb = excel.Workbooks.Open(tgt_path)
        excel.Calculation = -4135

        all_sheets = ['债元', '金元', '地元', '票', '融', '企', '信元', '其元']
        for name in all_sheets:
            if name in [s.Name for s in tgt_wb.Sheets]:
                ws = tgt_wb.Sheets(name)
                for i in range(n):
                    ws.Rows(3).Delete()
                print(f"  ✓ {name}：删除 {n} 行")

        tgt_wb.SaveAs(tgt_path + '.tmp.xlsx')
        print(f"\n回滚完成！")

    finally:
        tgt_wb.Close(SaveChanges=False)
        excel.Quit()
        pythoncom.CoUninitialize()
        time.sleep(3)
        tmp_path = tgt_path + '.tmp.xlsx'
        if os.path.exists(tmp_path):
            if os.path.exists(tgt_path):
                os.remove(tgt_path)
            os.rename(tmp_path, tgt_path)


def export_data():
    """运行数据导出"""
    print("\n导出 data.json...")
    import subprocess
    result = subprocess.run([sys.executable, 'export_data.py'],
                          capture_output=True, text=True, encoding='utf-8', errors='replace', cwd=os.getcwd())
    print(result.stdout)
    if result.stderr:
        print("STDERR:", result.stderr)
    if result.returncode == 0:
        print("✓ data.json 导出完成")
    else:
        print("⚠ 导出失败")
    return result.returncode == 0


def update_ultra_long():
    """超长债日报更新：下载邮件 → 按序补齐 → 导出"""
    print(f"\n{'=' * 60}")
    print("  超长利率债数据更新")
    print(f"{'=' * 60}")

    # 1. 从邮箱下载缺失的日报
    try:
        from fetch_email import fetch_ultra_long_daily
        downloaded = fetch_ultra_long_daily()
    except Exception as e:
        print(f"  邮箱下载跳过: {e}")
        downloaded = []

    # 2. 查找ultra_long目录下所有日报文件
    ultra_dir = os.path.join(os.getcwd(), 'ultra_long')
    pattern = re.compile(r'现券交易日报(\d{4})\.xlsx$')
    daily_files = []
    for f in os.listdir(ultra_dir):
        m = pattern.match(f)
        if m:
            daily_files.append((m.group(1), f))
    daily_files.sort(key=lambda x: x[0])  # 按日期从旧到新

    if not daily_files:
        print("  无日报文件可处理")
        return False

    # 2.5 更新前备份三个数据文件（防止写入中断导致损坏）
    bak_files = ['超长利率债买入数据.xlsx', '超长利率债卖出数据.xlsx', '超长利率债净买入数据.xlsx']
    for fname in bak_files:
        src = os.path.join(ultra_dir, fname)
        if os.path.exists(src):
            dst = os.path.join(ultra_dir, fname + '.bak')
            shutil.copy2(src, dst)
    print(f"  ✓ 已备份 {len([f for f in bak_files if os.path.exists(os.path.join(ultra_dir, f))])} 个数据文件")

    # 3. 逐个运行update_ultra_long.py（幂等：已存在的日期自动跳过）
    updated = 0
    for mmdd, filename in daily_files:
        print(f"\n  处理 {filename}...")
        result = subprocess.run(
            [sys.executable, 'update_ultra_long.py', filename],
            capture_output=True, text=True, encoding='utf-8', errors='replace',
            cwd=ultra_dir
        )
        if result.stdout:
            # 只打印关键行
            for line in result.stdout.strip().split('\n'):
                if '✓' in line or '已存在' in line or '更新完成' in line or '错误' in line.lower():
                    print(f"    {line.strip()}")
        if '已存在，跳过' not in (result.stdout or ''):
            updated += 1

    # 清理已处理的日报文件
    deleted = 0
    for mmdd, filename in daily_files:
        filepath = os.path.join(ultra_dir, filename)
        try:
            os.remove(filepath)
            deleted += 1
        except Exception as e:
            print(f"  ⚠ 无法删除 {filename}: {e}")
    if deleted:
        print(f"  ✓ 已清理 {deleted} 个日报文件")

    if updated == 0:
        print("\n  超长债数据已是最新")
        return False

    # 4. 导出 ultra_long_data.json
    print(f"\n  导出 ultra_long_data.json...")
    result = subprocess.run(
        [sys.executable, 'export_ultra_long.py'],
        capture_output=True, text=True, encoding='utf-8', errors='replace',
        cwd=ultra_dir
    )
    if result.stdout:
        for line in result.stdout.strip().split('\n'):
            if '完成' in line or '已保存' in line or '耗时' in line:
                print(f"    {line.strip()}")
    if result.returncode == 0:
        print("  ✓ ultra_long_data.json 导出完成")
    else:
        print(f"  ⚠ 导出失败: {(result.stderr or '')[:200]}")
        return False

    return True


def deploy(files):
    """部署文件到GitHub Pages"""
    deploy_log = os.path.join(os.getcwd(), 'deploy_output.tmp')
    cmd = [sys.executable, 'github_deploy.py', '--files'] + files
    with open(deploy_log, 'w', encoding='utf-8') as f:
        deploy_result = subprocess.run(
            cmd, stdout=f, stderr=subprocess.STDOUT, cwd=os.getcwd()
        )
    deploy_output = ''
    try:
        with open(deploy_log, 'r', encoding='utf-8') as f:
            deploy_output = f.read()
        os.remove(deploy_log)
    except Exception:
        pass
    print(deploy_output)
    return deploy_result.returncode == 0, deploy_output


def main():
    if '--rollback' in sys.argv:
        idx = sys.argv.index('--rollback')
        n = int(sys.argv[idx + 1]) if idx + 1 < len(sys.argv) else None
        if n is None:
            n = int(input("上次插入了几行？").strip())
        rollback(n)
        return

    # === 第0步：从QQ邮箱自动下载最新源文件 ===
    print("正在检查QQ邮箱是否有新的源文件...")
    try:
        from fetch_email import main as fetch_mail
        downloaded = fetch_mail()
        if downloaded:
            print(f"✓ 已从邮箱下载: {os.path.basename(downloaded)}")
        else:
            print("  邮箱中无新文件，使用本地已有文件")
    except Exception as e:
        print(f"  邮箱下载跳过: {e}")

    # 正常更新流程
    source_file = find_source_file()
    result = update(TARGET_FILE, source_file)
    n = result[0] if isinstance(result, tuple) else result

    # 超长债更新
    ultra_updated = False
    try:
        ultra_updated = update_ultra_long()
    except Exception as e:
        print(f"  ⚠ 超长债更新失败: {e}")

    if n > 0 or ultra_updated:
        if n > 0:
            print("等待Excel释放文件...")
            time.sleep(3)

        # 导出
        export_ok = False
        if n > 0:
            export_ok = export_data()

        # 部署
        deploy_files = []
        if export_ok:
            deploy_files.extend(['index.html', 'data.json'])
        if ultra_updated and os.path.exists('ultra_long_data.json'):
            if 'index.html' not in deploy_files:
                deploy_files.append('index.html')
            deploy_files.append('ultra_long_data.json')

        if deploy_files:
            print(f"\n{'=' * 60}")
            print("正在部署到GitHub Pages...")
            ok, output = deploy(deploy_files)
            if ok:
                print("✓ GitHub Pages 部署完成")
                print("  https://yhfund1.github.io/jgxw/")
                parts = []
                if n > 0:
                    parts.append(f"机构行为数据插入{n}行")
                if ultra_updated:
                    parts.append("超长债数据已更新")
                send_notification(
                    f"债券数据更新成功 - {', '.join(parts)}",
                    f"时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f"更新内容: {', '.join(parts)}\n"
                    f"GitHub Pages: https://yhfund1.github.io/jgxw/"
                )
            else:
                print("⚠ GitHub Pages 部署失败")
                send_notification(
                    f"债券数据更新但部署失败",
                    f"时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
                    f"部署失败原因: {output[:200] if output else '未知'}"
                )
            print(f"{'=' * 60}")
    else:
        print("✓ 数据已是最新，无需更新")
        # 确保无残留Excel进程（防止下次触发初始化失败）
        try:
            subprocess.run(['taskkill', '/IM', 'EXCEL.EXE', '/F'],
                           capture_output=True, timeout=10)
        except Exception:
            pass


if __name__ == '__main__':
    try:
        main()
    except Exception as e:
        import traceback
        err_msg = traceback.format_exc()
        print(f"\n✗ 脚本执行失败:\n{err_msg}")
        send_notification(
            f"债券数据更新失败 - {type(e).__name__}",
            f"时间: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
            f"错误类型: {type(e).__name__}\n"
            f"错误信息: {str(e)}\n\n"
            f"详细堆栈:\n{err_msg[:500]}"
        )

