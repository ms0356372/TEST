"""程式入口；支援無 GUI 的 --self-test。"""
from __future__ import annotations
import argparse
from pathlib import Path
from self_check import run_self_check, write_result_file


def main() -> int:
    parser=argparse.ArgumentParser();parser.add_argument("--self-test",action="store_true");parser.add_argument("--result-file",type=Path);args=parser.parse_args()
    if args.self_test:
        results=run_self_check();failed=any(x["status"]=="FAIL" for x in results)
        lines=[f"[{x['status']}] {x['name']}: {x['message']}" for x in results]
        print("SELF TEST\n"+"\n".join(lines)+f"\nRESULT: {'FAIL' if failed else 'PASS'}\n{len(results)-sum(x['status']=='FAIL' for x in results)} / {len(results)} checks passed")
        if args.result_file:write_result_file(args.result_file,results)
        return int(failed)
    from gui import launch_with_splash
    launch_with_splash();return 0

if __name__=="__main__":raise SystemExit(main())
