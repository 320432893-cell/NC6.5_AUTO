# 职责：判定 NC 是否被上游(finance 桌面版)当引擎子进程驱动(engine 模式),
#       供入口层据此抑制面向人工操作员的旁白/确认/交互打印,让 stdout 在引擎模式下
#       只剩{机器结果信封 + 必要进度};开发细节改走日志文件(见 core/logger)。
# 不做什么：不做日志配置、不解析业务、不碰 JAB。
# 允许依赖层：仅标准库(os)。
# 谁不应该 import：—(底座工具,任何层可 import)。
import os

ENGINE_MODE_ENV = "NC_ENGINE_MODE"


def is_engine_mode() -> bool:
    """NC 被 finance 桌面版当引擎子进程拉起时为 True(nc_engine 注入 NC_ENGINE_MODE=1)。

    引擎模式下入口层应抑制面向人工操作员的旁白/确认/交互(input)——用户进度走
    run_state.json、结果走机器信封;CLI 直跑(无此环境)保持原有旁白与确认。
    """
    return os.environ.get(ENGINE_MODE_ENV, "").strip().lower() in {"1", "true", "yes"}
