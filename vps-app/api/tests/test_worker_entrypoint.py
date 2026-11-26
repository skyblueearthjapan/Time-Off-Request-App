"""`python -m app.worker` で起動したとき、各担当の予定が実際に回る一覧に入ること（2026-11-26 の不具合の再発防止）。"""
import runpy


def test_module_entrypoint_delegates_to_app_worker_main(monkeypatch):
    from app import worker

    called = {}

    def fake_main(argv=None):
        called["module"] = worker.__name__
        called["schedules"] = sorted(worker._schedules)

    monkeypatch.setattr(worker, "main", fake_main)
    runpy.run_module("app.worker", run_name="__main__")
    assert called.get("module") == "app.worker"


def test_registrations_land_in_app_worker_schedules(monkeypatch):
    """担当 D・M の register_worker() が登録した予定は、app.worker._schedules に入る。"""
    import sys
    import types

    from app import worker

    module = types.ModuleType("app.services.summary_sync")

    def register_worker():
        from app import worker as w

        w.register_schedule("rebuild_leave_summary", task=lambda sf: {"ok": True}, at=(6, 50))

    module.register_worker = register_worker
    monkeypatch.setitem(sys.modules, "app.services.summary_sync", module)
    assert "app.services.summary_sync" in worker.load_registrations()
    assert "rebuild_leave_summary" in worker._schedules


def test_registration_modules_have_no_removed_features():
    from app import worker

    joined = " ".join(worker.REGISTRATION_MODULES)
    for removed in ("morning_batch", "sheet_intake"):
        assert removed not in joined
