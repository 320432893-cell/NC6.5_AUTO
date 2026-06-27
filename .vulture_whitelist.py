# vulture 白名单基线 —— check.py deep 的 vulture 闸读它,名单内不报。
# 作用:**新出现的死函数/死码(不在此名单)才会让闸变红**,实现「公开零调用」机器闸(补缝三)。
# 两类:1) 预留 API/测试桩(初步稳定期保留,定型后回头复核删留);
#       2) vulture 假阳性(dataclass 字段经序列化用、openpyxl 属性、__exit__ 协议参数——静态看不见)。
# 纪律:加条目=一次放行登记;真死码直接删勿塞此。重生成:
#   vulture core tools tests --exclude '*/archive/*' --min-confidence 60 --make-whitelist
# 基线已批:2026-06-27(用户"继续排除"全收,含 8 个预留API函数);此后新增条目=待批 pending,不自动生效。

_.collect_page_controls  # unused method (core/jab_batch_processor.py:192)
_.collect_window_controls  # unused method (core/jab_batch_processor.py:197)
_.take_screenshot  # unused method (core/jab_operator.py:199)
voucher_text  # unused variable (core/models.py:47)
extra_text  # unused variable (core/models.py:48)
fallback_reason  # unused variable (core/models.py:82)
_.build_report  # unused method (core/nc_page_probe.py:25)
_.schema_version  # unused attribute (core/receipt_config.py:24)
display_name  # unused variable (core/receipt_models.py:35)
nc_done_status  # unused variable (core/receipt_models.py:70)
_.font  # unused attribute (core/receipt_sheet.py:218)
_.font  # unused attribute (core/receipt_sheet.py:225)
created_at  # unused variable (core/voucher_plan_cache.py:25)
exc_type  # unused variable (tests/test_nc_pending_workflow.py:15)
exc_type  # unused variable (tests/test_nc_switch_generated_workflow.py:8)
exc_type  # unused variable (tests/test_nc_table_matcher.py:36)
no_real_keyboard_abort  # unused function (tests/test_nc_voucher_workflow.py:10)
exc_type  # unused variable (tests/test_nc_voucher_workflow.py:19)
exc_type  # unused variable (tests/test_receipt_detail_async_verifier.py:24)
default_counterparty_header_ok  # unused function (tests/test_receipt_full_flow_entry.py:125)
_.wait_save_success  # unused method (tests/test_receipt_full_flow_entry.py:540)
_.wait_calls  # unused attribute (tests/test_receipt_query_fill.py:52)
_.wait_calls  # unused attribute (tests/test_receipt_query_fill.py:63)
a  # unused variable (tests/test_receipt_self_made_fill_trial.py:161)
k  # unused variable (tests/test_receipt_self_made_fill_trial.py:161)
FakePreload  # unused class (tests/test_receipt_self_made_fill_trial.py:1870)
_.finished_at  # unused attribute (tools/receipt_detail_async_verifier.py:38)
_.finished_at  # unused attribute (tools/receipt_detail_async_verifier.py:193)
BUSINESS_TYPE_COL  # unused variable (tools/receipt_detail_fields.py:128)
read_first_row_cells  # unused function (tools/receipt_detail_reader.py:62)
get_cell_context  # unused function (tools/receipt_detail_screen_writer.py:54)
find_query_condition_scope_path_only  # unused function (tools/receipt_query_dynamic_fields.py:69)
field_cfg  # unused variable (tools/receipt_query_fill.py:110)
get_window_pid  # unused function (tools/receipt_self_made_fill_trial.py:1324)
foreground_root_hwnd  # unused function (tools/receipt_self_made_fill_trial.py:2416)
