from tools.jab_probe import JOBJECT, enum_windows


WATCH_NAMES = (
    "单据生成",
    "查询",
    "生成",
    "前台生成",
    "正式单据",
    "确定",
    "取消",
    "制单",
)


class NCPageProbe:
    def __init__(self, jab, batch_cfg):
        self.jab = jab
        self.batch_cfg = batch_cfg

    def build_report(self, max_rows=5, max_cols=25):
        self.jab.ensure_started()
        windows = self.collect_java_windows()
        controls = self.collect_controls()
        tables = self.jab.read_all_table_cells(max_rows=max_rows, max_cols=max_cols)

        voucher_window = self.batch_cfg.get("voucher_window_title", "制单")
        query_title = self.batch_cfg.get("open_query", {}).get(
            "dialog_title", "查询"
        )
        blockers = [
            window
            for window in windows
            if window["title"] in (voucher_window, query_title)
            and window["class"] == "SunAwtDialog"
            and window["visible"]
        ]

        return {
            "blocking_child_windows": blockers,
            "parent_markers": [
                item for item in controls if item["name"] == "单据生成"
            ],
            "watched_controls": controls,
            "table_signatures": [
                self.describe_table(table) for table in tables
            ],
        }

    def collect_java_windows(self):
        windows = []
        for hwnd, title, class_name, pid, visible in enum_windows(
            include_children=True
        ):
            try:
                is_java = bool(self.jab.dll.isJavaWindow(hwnd))
            except Exception:
                is_java = False
            if is_java:
                windows.append(
                    {
                        "hwnd": int(hwnd),
                        "title": title,
                        "class": class_name,
                        "pid": pid,
                        "visible": bool(visible),
                    }
                )
        return windows

    def collect_controls(self):
        found = []
        seen = set()
        for hwnd, title, class_name, pid, visible in enum_windows(
            include_children=True
        ):
            if not visible or not self.jab.dll.isJavaWindow(hwnd):
                continue
            vm_id, root = self.get_root(hwnd)
            if root is None:
                continue
            self.collect_controls_in_tree(
                vm_id,
                root,
                [],
                {
                    "hwnd": int(hwnd),
                    "title": title,
                    "class": class_name,
                    "pid": pid,
                    "visible": bool(visible),
                },
                found,
                seen,
                0,
            )
        return found

    def get_root(self, hwnd):
        import ctypes

        vm_id = ctypes.c_long()
        root_context = JOBJECT()
        if not self.jab.dll.getAccessibleContextFromHWND(
            hwnd,
            ctypes.byref(vm_id),
            ctypes.byref(root_context),
        ):
            return None, None
        return vm_id.value, root_context.value

    def collect_controls_in_tree(
        self,
        vm_id,
        context,
        path,
        window,
        found,
        seen,
        depth,
    ):
        info = self.jab.get_context_info(vm_id, context)
        if not info:
            return

        name = info.name.strip()
        desc = info.description.strip()
        role = info.role_en_US.strip() or info.role.strip()
        states = info.states_en_US.strip() or info.states.strip()
        role_l = role.lower()
        states_l = states.lower()
        text = name or desc

        if text in WATCH_NAMES or any(marker in text for marker in WATCH_NAMES):
            key = (window["hwnd"], ".".join(map(str, path)), name, desc, role)
            if key not in seen:
                seen.add(key)
                found.append(
                    {
                        "window_title": window["title"],
                        "window_class": window["class"],
                        "path": ".".join(map(str, path)),
                        "name": name,
                        "description": desc,
                        "role": role,
                        "states": states,
                        "showing": "visible" in states_l
                        and "showing" in states_l,
                        "bounds": [info.x, info.y, info.width, info.height],
                    }
                )

        if depth >= self.jab.max_depth or role_l == "table":
            return

        child_count = min(info.childrenCount, self.jab.max_children)
        for index in range(child_count):
            child = self.jab.dll.getAccessibleChildFromContext(vm_id, context, index)
            if not child:
                continue
            self.collect_controls_in_tree(
                vm_id, child, path + [index], window, found, seen, depth + 1
            )
            self.jab.release_contexts(vm_id, [child])

    def describe_table(self, table):
        generated_date_col = self.batch_cfg.get("generated_date_col", 18)
        voucher_col = self.batch_cfg.get("generated_voucher_col", 22)
        return {
            "table_index": table["table_index"],
            "window_title": table.get("window_title"),
            "window_class": table.get("window_class"),
            "row_count": table["row_count"],
            "col_count": table["col_count"],
            "date_col": generated_date_col,
            "date_values": sample_col(table, generated_date_col)[:8],
            "voucher_col": voucher_col,
            "voucher_values": sample_col(table, voucher_col)[:8],
            "sample_rows": [
                {
                    "row_index": row["row_index"],
                    "cells": row["cells"],
                    "selected": row["selected"],
                }
                for row in table["rows"]
            ],
        }

    def collect_named_controls(
        self,
        names,
        window_title=None,
        window_class=None,
        require_showing=True,
    ):
        self.jab.ensure_started()
        controls = []
        seen = set()
        for name in names:
            context, vm_id, owned, path = self.jab.find_context_once_with_path(
                name,
                normalized_roles=[],
                require_showing=require_showing,
                window_title=window_title,
                window_class=window_class,
                visible_only=True,
            )
            if not context:
                continue
            info = self.jab.get_context_info(vm_id, context)
            if not info:
                self.jab.release_contexts(vm_id, owned)
                continue
            states = (info.states_en_US.strip() or info.states.strip()).lower()
            key = (name, ".".join(map(str, path)))
            if key not in seen:
                seen.add(key)
                controls.append(
                    {
                        "name": name,
                        "path": ".".join(map(str, path)),
                        "role": info.role_en_US.strip() or info.role.strip(),
                        "showing": "visible" in states and "showing" in states,
                    }
                )
            self.jab.release_contexts(vm_id, owned)
        return controls

    def collect_visible_buttons_by_desc_tokens(
        self,
        tokens,
        window_title=None,
        window_class=None,
    ):
        self.jab.ensure_started()
        found = []
        for hwnd, title, class_name, _pid, visible in enum_windows(
            include_children=True
        ):
            if window_title is not None and title != window_title:
                continue
            if window_class is not None and class_name != window_class:
                continue
            if not visible or not self.jab.dll.isJavaWindow(hwnd):
                continue
            vm_id, root = self.get_root(hwnd)
            if root is None:
                continue
            self.collect_buttons_by_desc_in_tree(
                vm_id,
                root,
                [],
                {"title": title, "class": class_name},
                set(tokens),
                found,
                0,
            )
        return found

    def collect_buttons_by_desc_in_tree(
        self,
        vm_id,
        context,
        path,
        window,
        tokens,
        found,
        depth,
    ):
        info = self.jab.get_context_info(vm_id, context)
        if not info:
            return
        role = (info.role_en_US.strip() or info.role.strip()).lower()
        states = (info.states_en_US.strip() or info.states.strip()).lower()
        desc = info.description.strip()

        if (
            "button" in role
            and "visible" in states
            and "showing" in states
            and any(token in desc for token in tokens)
        ):
            found.append(
                {
                    "token": next(token for token in tokens if token in desc),
                    "path": ".".join(map(str, path)),
                    "name": info.name.strip(),
                    "description": desc,
                    "role": role,
                    "showing": True,
                    "window_title": window["title"],
                    "window_class": window["class"],
                }
            )

        if depth >= self.jab.max_depth or role == "table":
            return
        child_count = min(info.childrenCount, self.jab.max_children)
        for index in range(child_count):
            child = self.jab.dll.getAccessibleChildFromContext(vm_id, context, index)
            if not child:
                continue
            self.collect_buttons_by_desc_in_tree(
                vm_id, child, path + [index], window, tokens, found, depth + 1
            )
            self.jab.release_contexts(vm_id, [child])

    def read_page_table_signatures(
        self,
        generated_date_col,
        voucher_col,
        amount_col,
        partner_col,
    ):
        tables = self.jab.read_all_table_cells(max_rows=25, max_cols=30)
        return [
            describe_signature_table(
                self.jab,
                table,
                generated_date_col,
                voucher_col,
                amount_col,
                partner_col,
            )
            for table in tables
        ]


def describe_signature_table(
    jab,
    table,
    generated_date_col,
    voucher_col,
    amount_col,
    partner_col,
):
    return {
        "table_index": table["table_index"],
        "window_title": table.get("window_title"),
        "window_class": table.get("window_class"),
        "row_count": table["row_count"],
        "col_count": table["col_count"],
        "date_values": sample_col(table, generated_date_col),
        "voucher_values": sample_col(table, voucher_col),
        "rows": [
            {
                "row_index": row["row_index"],
                "amount": jab.normalize_amount(
                    row["cells"][amount_col] if amount_col < len(row["cells"]) else ""
                ),
                "partner": jab.normalize_text(
                    row["cells"][partner_col]
                    if partner_col < len(row["cells"])
                    else ""
                ),
            }
            for row in table.get("rows", [])
        ],
    }


def sample_col(table, col):
    values = []
    if col is None:
        return values
    for row in table.get("rows", []):
        cells = row.get("cells", [])
        if 0 <= col < len(cells):
            text = str(cells[col]).strip()
            if text:
                values.append(text)
    return values
