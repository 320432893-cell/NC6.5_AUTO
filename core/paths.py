# 职责: 解析运行时基目录与可写产物目录(日志/状态/停止标志),屏蔽"源码态 vs 冻结打包态"与外部注入差异。
# 不做什么: 不读业务配置;不写业务数据;不碰 JAB/界面;不定义日志格式(那是 logger.py)。
# 允许依赖层: 仅标准库(os/sys/pathlib)。
# 谁不应该 import: 无层级限制(基础层);但本模块不得 import core 内其它模块,以免环。
import os
import sys
from pathlib import Path

# GUI 外壳/控制进程可注入的覆盖项(见 ENGINE_CONTRACT.md §1.5/§1.6)。
RUNTIME_DIR_ENV = "NC_RUNTIME_DIR"
STOP_FLAG_NAME = "abort.flag"


def base_dir() -> Path:
    """安装/源码基目录。冻结打包态用可执行文件所在目录,源码态用仓库根。"""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    # core/paths.py -> parents[0]=core/ , parents[1]=仓库根
    return Path(__file__).resolve().parents[1]


def runtime_dir() -> Path:
    """可写运行时根。外部进程可用 NC_RUNTIME_DIR 注入(打包后基目录可能只读)。"""
    override = os.environ.get(RUNTIME_DIR_ENV)
    if override:
        return Path(override)
    return base_dir()


def logs_dir(create: bool = True) -> Path:
    """日志/状态/性能等可写产物目录。热路径(只读探测)传 create=False 免每次建目录。"""
    target = runtime_dir() / "logs"
    if create:
        target.mkdir(parents=True, exist_ok=True)
    return target


def stop_flag_path() -> Path:
    """外部停止标志文件路径。控制进程创建它即请求引擎在下一个安全检查点停止。"""
    return logs_dir(create=False) / STOP_FLAG_NAME


def clear_stop_flag() -> None:
    """运行开始时清掉残留标志,避免上一次的停止请求误杀本次。"""
    try:
        stop_flag_path().unlink(missing_ok=True)
    except OSError:
        pass
