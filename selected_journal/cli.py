"""Command-line entry point."""
import argparse
import json
import sys
from datetime import datetime
from pathlib import Path

from .client import LetPubClient, atomic_json
from .inputs import read_file, read_lines
from .runner import run_batch
from .storage import read_events


def parser():
    root = argparse.ArgumentParser(description="LetPub 期刊名称 / ISSN 批量查询")
    commands = root.add_subparsers(dest="command", required=True)
    query = commands.add_parser("query", help="查询期刊并导出 Excel / CSV / JSON")
    source = query.add_mutually_exclusive_group(required=True)
    source.add_argument("--input", type=Path, help="TXT / CSV / XLSX 名单")
    source.add_argument("--stdin", action="store_true", help="从标准输入读取，每行一个名称或 ISSN")
    source.add_argument("--resume", type=Path, help="恢复已有任务目录")
    query.add_argument("--output", type=Path, help="新任务输出目录；不与 --resume 同用")
    query.add_argument("--cookie-file", type=Path, help="可选 Netscape 格式 LetPub Cookie 文件")
    query.add_argument("--refresh", action="store_true", help="不读取 24 小时详情缓存")
    return root


def main(argv=None):
    argument_parser = parser()
    try:
        args = argument_parser.parse_args(argv)
    except SystemExit as error:
        return 0 if error.code == 0 else 1
    client = None
    try:
        if args.resume and args.output:
            raise ValueError("--resume 不与 --output 同用")
        if args.resume:
            directory = args.resume.resolve()
            manifest = json.loads((directory / "inputs.json").read_text(encoding="utf-8"))
            if manifest.get("schema_version") != 1 or not manifest.get("inputs"):
                raise ValueError("任务输入快照格式无效")
            inputs = manifest["inputs"]
            events = read_events(directory)
        else:
            if args.stdin:
                if sys.stdin.isatty():
                    print("粘贴期刊名称或 ISSN，每行一个；Ctrl+D 结束输入。", flush=True)
                inputs = read_lines(sys.stdin)
            else:
                inputs = read_file(args.input)
            if not inputs:
                raise ValueError("名单为空，没有可处理的期刊")
            directory = (args.output or Path("runs") / datetime.now().astimezone().strftime("%Y%m%d-%H%M%S-%f")).resolve()
            if directory.exists() and any(directory.iterdir()):
                raise ValueError("输出目录非空，请指定新目录或使用 --resume")
            events = {}
        # Resolve credentials before creating a new task so configuration errors
        # don't leave a misleading input snapshot.
        client = LetPubClient(directory.parent / ".letpub-cache", args.cookie_file, args.refresh)
        if not args.resume:
            directory.mkdir(parents=True, exist_ok=True)
            atomic_json(directory / "inputs.json", {"schema_version": 1, "inputs": inputs})
        print(f"任务目录: {directory}\n访问模式: {'已加载本地 Cookie（是否登录由页面决定）' if client.mode != 'anonymous' else '匿名'}", flush=True)
        code, _ = run_batch(inputs, directory, client, events, emit=lambda message: print(message, flush=True))
        print(f"结果: {directory / 'results.xlsx'}", flush=True)
        return code
    except KeyboardInterrupt:
        return 130
    except (ValueError, OSError, KeyError, TypeError) as error:
        print(f"配置或保存错误: {error}", file=sys.stderr)
        return 1
    finally:
        if client is not None:
            client.close()


if __name__ == "__main__":
    raise SystemExit(main())
