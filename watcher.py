# -*- coding: utf-8 -*-
"""
菌种分布图 HTML 自动更新监视器
=================================

功能：
    定期检查 细菌库.xlsx 文件是否发生变化（新增 / 修改 / 被覆盖）。
    一旦检测到变化，自动调用 generate_dashboard.py 重新生成菌种分布图 HTML。

使用方式：
    python watcher.py                     # 默认每小时检查一次
    python watcher.py --interval 300      # 每 300 秒（5 分钟）检查一次
    python watcher.py --once              # 只检查并生成一次后立即退出

原理说明：
    金山文档（kdocs.cn / WPS）没有对个人开放的公开数据读取 API，
    因此需要你把金山文档里的表格手动「导出为 xlsx」
    并覆盖本目录下的 细菌库.xlsx 文件。
    本监视器会在后台持续监听这个本地文件，一旦它被更新，
    就自动重新生成 HTML 报告。

    你只需要做一件事：
    在金山文档里改完数据 → 导出 xlsx → 覆盖本目录的 细菌库.xlsx
    剩下的（重新生成 HTML）交给本脚本自动完成。
"""

import os
import sys
import time
import json
import hashlib
import subprocess
from datetime import datetime

# ============================================================
# 配置
# ============================================================

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 要监视的数据文件（你需要从这里手动覆盖更新金山导出的 xlsx）
DATA_FILE = os.path.join(BASE_DIR, "细菌库.xlsx")

# 生成 HTML 的脚本
GENERATOR_SCRIPT = os.path.join(BASE_DIR, "generate_dashboard.py")

# Python 解释器（你的 Anaconda 环境）
PYTHON_EXE = r"C:\Users\Administrator\anaconda3\python.exe"

# 状态文件（记录上次检测到的文件指纹）
STATE_FILE = os.path.join(BASE_DIR, ".watcher_state.json")

# 日志文件
LOG_FILE = os.path.join(BASE_DIR, "watcher.log")


# ============================================================
# 小工具函数
# ============================================================

def log(message, to_console=True):
    """写日志：同时输出到控制台和日志文件。"""
    line = f"[{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}] {message}"
    if to_console:
        print(line)
    try:
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass


def get_file_fingerprint(path):
    """
    计算文件指纹（内容哈希 + 大小 + 修改时间）。
    只要文件内容或大小发生变化，指纹就会不同。
    """
    if not os.path.exists(path):
        return None

    h = hashlib.sha256()
    try:
        with open(path, "rb") as f:
            # 分块读取，避免大文件一次性占用过多内存
            for chunk in iter(lambda: f.read(1024 * 1024), b""):
                h.update(chunk)
    except Exception as e:
        log(f"警告：读取 {path} 失败：{e}")
        return None

    stat = os.stat(path)
    return {
        "sha256": h.hexdigest(),
        "size": stat.st_size,
        "mtime": stat.st_mtime,
    }


def load_state():
    """读取上次保存的文件指纹。"""
    if os.path.exists(STATE_FILE):
        try:
            with open(STATE_FILE, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return None
    return None


def save_state(state):
    """保存当前文件指纹。"""
    try:
        with open(STATE_FILE, "w", encoding="utf-8") as f:
            json.dump(state, f, ensure_ascii=False, indent=2)
    except Exception as e:
        log(f"警告：保存状态失败：{e}")


def run_generator():
    """
    调用 generate_dashboard.py 重新生成 HTML。
    返回 (成功与否, 输出文本)。
    """
    log("检测到数据变化，开始重新生成 HTML ...")
    try:
        proc = subprocess.run(
            [PYTHON_EXE, GENERATOR_SCRIPT],
            cwd=BASE_DIR,
            capture_output=True,
            text=True,
            encoding="utf-8",
            errors="replace",
            timeout=600,  # 10 分钟超时
        )

        output = (proc.stdout or "") + (proc.stderr or "")

        if proc.returncode == 0:
            log("✓ HTML 重新生成成功。")
            return True, output
        else:
            log(f"✗ 生成失败（退出码 {proc.returncode}）。")
            log("失败输出：\n" + output.strip()[-2000:])
            return False, output

    except subprocess.TimeoutExpired:
        log("✗ 生成超时（超过 10 分钟）。")
        return False, "超时"
    except FileNotFoundError:
        log(f"✗ 找不到 Python 解释器：{PYTHON_EXE}")
        log("    请检查 watcher.py 顶部的 PYTHON_EXE 路径是否正确。")
        return False, "Python 解释器不存在"
    except Exception as e:
        log(f"✗ 生成过程异常：{e}")
        return False, str(e)


def check_once(force_generate=False):
    """
    检查一次数据文件是否变化，变化则重新生成。

    返回 True 表示（本次无需处理或处理成功），
    返回 False 表示生成失败。
    """
    if not os.path.exists(DATA_FILE):
        log(f"✗ 未找到数据文件：{DATA_FILE}")
        log("    请确保 细菌库.xlsx 存在于本目录。")
        return False

    current = get_file_fingerprint(DATA_FILE)
    if current is None:
        return False

    previous = load_state()

    changed = False
    if previous is None:
        # 第一次运行：没有历史记录
        if force_generate:
            log("首次运行，强制生成一次 HTML。")
            changed = True
        else:
            log("首次运行，已记录当前文件指纹（下次变化时自动生成）。")
    else:
        if (
            previous.get("sha256") != current["sha256"]
            or previous.get("size") != current["size"]
        ):
            changed = True
            log(
                f"检测到 细菌库.xlsx 已更新 "
                f"(大小 {previous.get('size', '?')} → {current['size']} 字节)。"
            )
        else:
            log("细菌库.xlsx 无变化，无需重新生成。")

    # 无论是首次还是有变化，都更新状态文件
    save_state(current)

    if changed:
        ok, _ = run_generator()
        return ok
    return True


# ============================================================
# 主流程
# ============================================================

def main():
    # 解析命令行参数
    interval = 3600  # 默认 1 小时
    once = False
    force = False

    args = sys.argv[1:]
    i = 0
    while i < len(args):
        a = args[i]
        if a in ("--interval", "-i"):
            i += 1
            if i < len(args):
                interval = int(args[i])
        elif a == "--once":
            once = True
        elif a == "--force":
            force = True
        elif a in ("--help", "-h"):
            print(__doc__)
            return
        else:
            print(f"未知参数：{a}")
            print(__doc__)
            return
        i += 1

    if interval < 10:
        interval = 10

    log("=" * 60)
    log("菌种分布图 HTML 自动更新监视器 已启动")
    log(f"数据文件：{DATA_FILE}")
    log(f"生成脚本：{GENERATOR_SCRIPT}")
    log(f"检查间隔：{interval} 秒")
    log(f"运行模式：{'单次执行' if once else '持续监听'}")
    log("=" * 60)

    if once:
        check_once(force_generate=force)
        log("单次执行完成，程序退出。")
        return

    # 持续监听：启动时先强制检查一次（确保数据是最新的）
    log("提示：程序将持续运行，请勿关闭本窗口。按 Ctrl+C 停止。")
    log("你只需把金山文档导出为 细菌库.xlsx 并覆盖到本目录即可。")
    check_once(force_generate=False)

    try:
        while True:
            time.sleep(interval)
            check_once()
    except KeyboardInterrupt:
        log("收到停止信号，监视器已退出。")


if __name__ == "__main__":
    main()