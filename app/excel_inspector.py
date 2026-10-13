
from pathlib import Path
import sys
import win32com.client as win32


def inspect_workbook(file_path: str) -> None:
    """Read basic workbook information without saving changes."""

    path = Path(file_path).expanduser().resolve()

    if not path.is_file():
        print(f"ERROR: File not found: {path}")
        return

    excel = None
    workbook = None

    try:
        # Start a separate Excel instance for this task.
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False

        # Open read-only and do not update external links.
        workbook = excel.Workbooks.Open(
            Filename=str(path),
            UpdateLinks=0,
            ReadOnly=True,
            IgnoreReadOnlyRecommended=True,
            AddToMru=False,
        )

        print("\nWORKBOOK INSPECTION")
        print("=" * 50)
        print(f"File: {path.name}")
        print(f"Worksheets: {workbook.Worksheets.Count}")

        for index in range(1, workbook.Worksheets.Count + 1):
            sheet = workbook.Worksheets.Item(index)

            try:
                used_range = sheet.UsedRange

                print(f"\nWorksheet {index}: {sheet.Name}")
                print(f"  Used rows: {used_range.Rows.Count}")
                print(f"  Used columns: {used_range.Columns.Count}")

                # Display a small sample of non-empty cell values.
                max_rows = min(used_range.Rows.Count, 5)
                max_cols = min(used_range.Columns.Count, 6)

                print("  Sample values:")

                for row in range(1, max_rows + 1):
                    values = []

                    for col in range(1, max_cols + 1):
                        value = used_range.Cells(row, col).Value
                        values.append(repr(value))

                    print("   ", " | ".join(values))

            finally:
                # Release this worksheet reference.
                sheet = None

        print("\nInspection completed.")
        print("The workbook was opened read-only and was not saved.")

    except Exception as exc:
        print(f"\nERROR: Could not inspect workbook.\n{exc}")

    finally:
        if workbook is not None:
            try:
                workbook.Close(SaveChanges=False)
            except Exception:
                pass

        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass

        workbook = None
        excel = None


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Excel Inspector")
        print("-" * 40)
        print("Usage:")
        print('python app\\excel_inspector.py "C:\\path\\to\\workbook.xlsx"')
        print("\nUse a COPY of a workbook for your first test.")
    else:
        inspect_workbook(sys.argv[1])
