import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.jab_operator import JABOperator  # noqa: E402
from core.utils import load_config  # noqa: E402


def main():
    rows = [int(arg) for arg in sys.argv[1:]] or [0, 2]
    jab = JABOperator(load_config())
    jab.ensure_started()

    try:
        target = None
        for index, (context, vm_id, owned, info, window) in enumerate(
            jab.find_tables_once()
        ):
            if (
                window.get("title") == "制单"
                and info.rowCount > 1
                and info.columnCount == 13
            ):
                target = (index, context, vm_id, owned, info, window)
                break

        if not target:
            print("no voucher table")
            return 1

        table_index, context, vm_id, owned, info, window = target
        print("target_table", table_index, window, info.rowCount, info.columnCount)

        jab.dll.clearAccessibleSelectionFromContext(vm_id, context)
        for row in rows:
            child_index = row * info.columnCount
            jab.dll.addAccessibleSelectionFromContext(vm_id, context, child_index)

        selected = jab.get_selected_child_indexes(
            vm_id, context, info.rowCount * info.columnCount
        )
        print("selected_indexes", selected[:80])

        for row in range(min(info.rowCount, 10)):
            values = []
            selected_cells = []
            for col in range(min(info.columnCount, 6)):
                text, is_selected = jab.get_table_cell_text_and_selection(
                    vm_id, context, row, col
                )
                values.append(text)
                selected_cells.append(is_selected)
            print("row", row, "selected", any(selected_cells), values)

        jab.release_contexts(vm_id, owned)
        return 0
    finally:
        jab.close()


if __name__ == "__main__":
    raise SystemExit(main())
