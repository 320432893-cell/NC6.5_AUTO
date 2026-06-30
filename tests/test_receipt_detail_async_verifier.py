# 覆盖的业务场景：收款单明细后台 verifier 复用主流程 JAB，不再新建/关闭 JAB 会话

from core.receipt_detail_async_verifier import DetailPipelineVerifier


def test_pipeline_verifier_reuses_external_jab_without_lifecycle(monkeypatch):
    calls = {"created": 0, "ensure": 0, "close": 0, "preload": 0, "lock": 0}

    class ForbiddenJAB:
        def __init__(self, _config):
            calls["created"] += 1

    class ExternalJAB:
        def ensure_started(self):
            calls["ensure"] += 1

        def close(self):
            calls["close"] += 1

    class CountingLock:
        def __enter__(self):
            calls["lock"] += 1

        def __exit__(self, exc_type, exc, traceback):
            return False

    monkeypatch.setattr(
        "core.receipt_detail_async_verifier.JABOperator",
        ForbiddenJAB,
    )
    monkeypatch.setattr(
        "core.receipt_detail_async_verifier.locate_receipt_body_table_cached",
        lambda *_args, **_kwargs: (
            calls.__setitem__("preload", calls["preload"] + 1)
            or {"best": {"path": "0.1", "row_count": 1, "col_count": 25}}
        ),
    )

    verifier = DetailPipelineVerifier(
        {},
        {"best": {"path": "0.1", "window": {}}},
        jab=ExternalJAB(),
        jab_lock=CountingLock(),
    )
    verifier.start()
    verifier.close(timeout=1.0)

    assert calls == {"created": 0, "ensure": 0, "close": 0, "preload": 1, "lock": 1}
