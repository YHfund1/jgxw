"""
从QQ邮箱自动下载【现券市场交易情况汇总跟踪】Excel附件
使用IMAP协议连接QQ邮箱，搜索最新邮件并下载xlsx附件
"""
import imaplib
import email
from email.header import decode_header
import os
import re
import glob
import time

# === 配置 ===
EMAIL_ADDR = '906881211@qq.com'
AUTH_CODE = 'jfblwqnjdzgqbajh'  # QQ邮箱IMAP授权码
IMAP_SERVER = 'imap.qq.com'
IMAP_PORT = 993
DOWNLOAD_DIR = os.path.dirname(os.path.abspath(__file__))

# 邮件过滤条件
SENDER_EMAIL = 'ficcresearch@hfzq.com.cn'
SUBJECT_KEYWORD = '现券市场交易情况汇总跟踪'


def decode_mime_str(s):
    """解码邮件头中的编码字符串"""
    if s is None:
        return ''
    parts = decode_header(s)
    result = []
    for part, charset in parts:
        if isinstance(part, bytes):
            result.append(part.decode(charset or 'utf-8', errors='replace'))
        else:
            result.append(part)
    return ''.join(result)


def find_latest_email(mail):
    """
    搜索来自华福固收的最新邮件
    优先用 FROM + SINCE 过滤，再在Python中按主题精确匹配
    返回: (msg, subject) 或 (None, None)
    """
    import datetime
    since_date = (datetime.date.today() - datetime.timedelta(days=14)).strftime('%d-%b-%Y')

    # 用发件人+日期过滤（高效）
    status, data = mail.search(None, f'(FROM "{SENDER_EMAIL}" SINCE "{since_date}")')
    if status != 'OK' or not data[0]:
        print(f'  最近14天未收到来自 {SENDER_EMAIL} 的邮件')
        return None, None

    msg_ids = data[0].split()
    msg_ids.reverse()  # 最新的在前
    print(f'  来自华福固收的邮件共 {len(msg_ids)} 封')

    for msg_id in msg_ids[:10]:
        # 获取邮件头
        status, msg_data = mail.fetch(msg_id, '(RFC822.HEADER)')
        if status != 'OK':
            continue

        header_msg = email.message_from_bytes(msg_data[0][1])
        subject = decode_mime_str(header_msg.get('Subject', ''))

        # 精确匹配主题关键词
        if SUBJECT_KEYWORD not in subject:
            continue

        # 从主题中提取日期后缀（如——0707中的0707）
        date_match = re.search(r'\u2014\u2014(\d{4})$', subject)
        if date_match:
            date_suffix = date_match.group(1)
            # 检查本地是否已有包含该日期的文件
            existing = glob.glob(os.path.join(DOWNLOAD_DIR, f'*{date_suffix}*快照*.xlsx'))
            if existing:
                print(f'  本地已有包含 {date_suffix} 的源文件，跳过下载')
                return None, None

        # 获取完整邮件
        status2, msg_data2 = mail.fetch(msg_id, '(RFC822)')
        if status2 != 'OK':
            continue

        raw_email = msg_data2[0][1]
        msg = email.message_from_bytes(raw_email)

        # 检查是否有xlsx附件
        has_xlsx = False
        for part in msg.walk():
            filename = part.get_filename()
            if filename and filename.endswith('.xlsx') and not filename.startswith('~$'):
                has_xlsx = True
                break

        if has_xlsx:
            return msg, subject
        else:
            print(f'  邮件 [{subject}] 无xlsx附件，跳过')

    return None, None


def download_attachments(msg, download_dir):
    """
    从邮件中提取xlsx附件并保存
    返回: 保存的文件路径列表
    """
    saved_files = []

    for part in msg.walk():
        filename = part.get_filename()
        if not filename:
            continue

        # 解码文件名
        filename = decode_mime_str(filename)

        # 只下载xlsx文件
        if not filename.endswith('.xlsx'):
            continue

        # 跳过临时文件
        if filename.startswith('~$'):
            continue

        filepath = os.path.join(download_dir, filename)
        content = part.get_payload(decode=True)

        with open(filepath, 'wb') as f:
            f.write(content)

        saved_files.append(filepath)
        print(f"  ✓ 已下载: {filename} ({len(content) / 1024 / 1024:.1f} MB)")

    return saved_files


def cleanup_old_sources(new_file):
    """删除旧的源文件，只保留最新下载的"""
    pattern = os.path.join(DOWNLOAD_DIR, '现券市场交易情况汇总跟踪*.xlsx')
    old_files = [f for f in glob.glob(pattern) if os.path.abspath(f) != os.path.abspath(new_file)]
    for f in old_files:
        try:
            os.remove(f)
            print(f"  ✓ 已删除旧文件: {os.path.basename(f)}")
        except Exception as e:
            print(f"  ⚠ 无法删除旧文件 {os.path.basename(f)}: {e}")


def main():
    """主函数：连接邮箱 → 搜索邮件 → 下载附件"""
    print("=" * 60)
    print("  QQ邮箱自动下载：现券市场交易情况汇总跟踪")
    print("=" * 60)

    # 连接IMAP服务器
    print(f"连接 {IMAP_SERVER}...")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ADDR, AUTH_CODE)
        print("  ✓ 登录成功")
    except Exception as e:
        print(f"  ✗ 登录失败: {e}")
        return None

    try:
        # 选择收件箱
        mail.select('INBOX')

        # 搜索最新邮件
        print(f"搜索来自华福固收的邮件...")
        msg, subject = find_latest_email(mail)

        if msg is None:
            print("  无新文件需要下载")
            return None

        print(f"  找到邮件: {subject}")

        # 检查邮件日期
        date_str = msg.get('Date', '')
        print(f"  邮件日期: {date_str}")

        # 下载附件
        print("下载附件...")
        saved_files = download_attachments(msg, DOWNLOAD_DIR)

        if not saved_files:
            print("  ✗ 邮件中没有xlsx附件")
            return None

        # 返回最新下载的文件路径
        latest_file = saved_files[0]
        print(f"\n最新文件: {os.path.basename(latest_file)}")

        # 清理旧的源文件（保留最新下载的，删除其他同模式文件）
        cleanup_old_sources(latest_file)

        return latest_file

    finally:
        try:
            mail.close()
            mail.logout()
        except:
            pass


# ============================================================
# 超长债日报批量下载（国海固收 颜子琦）
# ============================================================

ULTRA_LONG_SENDER = 'yanzq@188.com'
ULTRA_LONG_SUBJECT_KEYWORD = '现券交易'
ULTRA_LONG_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'ultra_long')


def fetch_ultra_long_daily():
    """
    从QQ邮箱批量下载所有缺失的【现券交易日报MMDD.xlsx】
    返回: 按日期排序的已下载文件路径列表（仅包含本次新下载的）
    """
    import datetime

    print("=" * 60)
    print("  QQ邮箱自动下载：现券交易日报（超长债）")
    print("=" * 60)

    # 连接IMAP
    print(f"连接 {IMAP_SERVER}...")
    try:
        mail = imaplib.IMAP4_SSL(IMAP_SERVER, IMAP_PORT)
        mail.login(EMAIL_ADDR, AUTH_CODE)
        print("  ✓ 登录成功")
    except Exception as e:
        print(f"  ✗ 登录失败: {e}")
        return []

    try:
        mail.select('INBOX')

        since_date = (datetime.date.today() - datetime.timedelta(days=21)).strftime('%d-%b-%Y')
        status, data = mail.search(None, f'(FROM "{ULTRA_LONG_SENDER}" SINCE "{since_date}")')
        if status != 'OK' or not data[0]:
            print(f'  最近21天未收到来自 {ULTRA_LONG_SENDER} 的邮件')
            return []

        msg_ids = data[0].split()
        print(f'  来自国海固收的邮件共 {len(msg_ids)} 封')

        # 收集所有匹配邮件: (date_mmdd, msg_id)
        candidates = []
        for msg_id in msg_ids:
            status, msg_data = mail.fetch(msg_id, '(RFC822.HEADER)')
            if status != 'OK':
                continue
            header_msg = email.message_from_bytes(msg_data[0][1])
            subject = decode_mime_str(header_msg.get('Subject', ''))

            if ULTRA_LONG_SUBJECT_KEYWORD not in subject:
                continue

            # 提取日期: "现券交易0710（...）" → "0710"
            date_match = re.search(r'现券交易(\d{4})', subject)
            if not date_match:
                continue

            mmdd = date_match.group(1)
            candidates.append((mmdd, msg_id, subject))

        if not candidates:
            print('  未找到现券交易日报邮件')
            return []

        # 按日期排序（从旧到新）
        candidates.sort(key=lambda x: x[0])

        # 检查本地已有哪些文件 + Excel中已有的日期
        os.makedirs(ULTRA_LONG_DIR, exist_ok=True)
        existing_pattern = re.compile(r'现券交易日报(\d{4})\.xlsx$')
        existing_dates = set()
        for f in os.listdir(ULTRA_LONG_DIR):
            m = existing_pattern.match(f)
            if m:
                existing_dates.add(m.group(1))

        # 从买入Excel读取已有日期，避免重复下载已入库的数据
        buy_file = os.path.join(ULTRA_LONG_DIR, '超长利率债买入数据.xlsx')
        if os.path.exists(buy_file):
            try:
                import openpyxl
                wb = openpyxl.load_workbook(buy_file, read_only=True, data_only=True)
                ws = wb['大型银行']
                for row in ws.iter_rows(min_row=2, max_col=1, values_only=True):
                    val = row[0]
                    if val is None:
                        continue
                    if hasattr(val, 'strftime') and val.year == datetime.date.today().year:
                        mmdd = val.strftime('%m%d')
                        existing_dates.add(mmdd)
                wb.close()
            except Exception:
                pass

        # 下载缺失的
        downloaded = []
        skipped = 0
        for mmdd, msg_id, subject in candidates:
            if mmdd in existing_dates:
                skipped += 1
                continue

            print(f"  下载 {mmdd}: {subject[:40]}...")
            status2, msg_data2 = mail.fetch(msg_id, '(RFC822)')
            if status2 != 'OK':
                continue

            raw_email = msg_data2[0][1]
            msg = email.message_from_bytes(raw_email)

            for part in msg.walk():
                filename = part.get_filename()
                if not filename:
                    continue
                filename = decode_mime_str(filename)
                if not filename.endswith('.xlsx') or filename.startswith('~$'):
                    continue

                filepath = os.path.join(ULTRA_LONG_DIR, filename)
                content = part.get_payload(decode=True)
                with open(filepath, 'wb') as f:
                    f.write(content)
                downloaded.append(filepath)
                print(f"    ✓ {filename} ({len(content)/1024/1024:.1f} MB)")

        if skipped:
            print(f"  已有 {skipped} 个文件跳过")
        if not downloaded:
            print("  无新文件需要下载")
        else:
            print(f"  ✓ 共下载 {len(downloaded)} 个文件")

        return downloaded

    finally:
        try:
            mail.close()
            mail.logout()
        except:
            pass


if __name__ == '__main__':
    import sys
    if len(sys.argv) > 1 and sys.argv[1] == '--ultra-long':
        result = fetch_ultra_long_daily()
        if result:
            print(f"\n✓ 完成！下载了 {len(result)} 个文件")
        else:
            print("\n无新文件")
    else:
        result = main()
        if result:
            print(f"\n✓ 完成！文件已保存到: {result}")
        else:
            print("\n✗ 下载失败")
