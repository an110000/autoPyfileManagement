"""主入口：无参数打开桌面窗口，--job 可在计划任务中重放导出的操作参数。"""

import argparse
import json
import os
import subprocess
import sys
from pathlib import Path


def desktop():
    """加载桌面程序；源码入口缺依赖时尝试项目解释器，提供明确安装指引。"""
    try:
        from batch_ui import main as start_ui
    except (ImportError, OSError) as error:
        folder = Path(__file__).resolve().parent
        if not getattr(sys, "frozen", False) and not os.environ.get("BATCH_UI_RETRY"):
            for name in (".venv-build", ".venv"):
                candidate = folder / name / "Scripts" / "python.exe"
                if candidate.is_file() and candidate.resolve() != Path(sys.executable).resolve():
                    env = dict(os.environ, BATCH_UI_RETRY="1")
                    return subprocess.call([str(candidate), str(folder / "main.py")], env=env)
        raise RuntimeError(f"桌面依赖加载失败（{sys.executable}）：{error}。请运行启动工作台.cmd，或用当前解释器安装 requirements.txt。") from error
    return start_ui()


def replay(args):
    """读取导出的清单或规则字典，执行同一批处理；可将机器可读结果写到指定路径。"""
    from batch_ops import perform

    config = json.loads(args.job.read_text(encoding="utf-8-sig"))
    options = config.get("options", config)
    if not isinstance(options, dict):
        raise ValueError("规则文件须为 JSON 对象")
    if "mode" not in options and "image_format" in options:
        options = dict(mode="一键归档", format=options["image_format"], sharpen=options["sharpen"])
    result, report = perform(args.input, args.output, options, preview=args.preview)
    report["result"] = str(result) if result else None
    if args.report:
        args.report.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    if sys.stdout:
        print(json.dumps(report["summary"], ensure_ascii=False))
    return 1 if report["summary"]["error"] else 0


def main():
    """选择图形或规则重放入口，记录启动错误；退出码 2 表示整批未能完成。"""
    parser = argparse.ArgumentParser(description="归档台桌面工作站 / 规则重放")
    parser.add_argument("--job", type=Path, help="导出的 JSON 清单或规则")
    parser.add_argument("--input", type=Path, help="重放时扫描的输入目录")
    parser.add_argument("--output", type=Path, help="正式成果的输出根目录")
    parser.add_argument("--report", type=Path, help="可选，单独保存本次 JSON 报告")
    parser.add_argument("--preview", action="store_true", help="预览本次规则和预期产物")
    args = parser.parse_args()
    try:
        if args.job is None:
            return desktop()
        if args.input is None or args.output is None:
            raise ValueError("--job 需要同时提供 --input 和 --output")
        return replay(args)
    except Exception as error:
        message = f"{type(error).__name__}：{error}"
        if args.report:
            args.report.write_text(json.dumps({"error": message}, ensure_ascii=False), encoding="utf-8")
        if sys.stderr:
            print(message, file=sys.stderr)
        elif not args.job and sys.platform == "win32":
            import ctypes
            ctypes.windll.user32.MessageBoxW(None, message, "归档台 · 启动错误", 0x10)
        return 2


if __name__ == "__main__":
    sys.exit(main())
