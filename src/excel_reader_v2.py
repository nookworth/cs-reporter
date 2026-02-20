# CS Reporter
# Excel Reader V2
# Excel file reader using pandas. This module is the bridge between raw Excel data and the PowerPoint template.
# **Date:** 2026-02-20
# **Status:** ✅ Completed
# **Component:** src

from pathlib import Path

from . import excel_utils_v2 as excel_utils 


class ExcelReader:

    def __init__(self, excel_path, config, previous_excel_path=None):
        self.excel_path = Path(excel_path)
        self.previous_excel_path = (
            Path(previous_excel_path) if previous_excel_path else None
        )
        self.config = config
        
        self.fields = config.get("fields", {})
        
        self.table_fields = config.get("table_fields", {})
        self.standard_excel_fields = config.get("standard_excel_fields", {})
        self.retail_excel_fields = config.get("retail_excel_fields", {})
        self.supplier_excel_fields = config.get("supplier_excel_fields", {})

    def extract_data(self):
        data = {}

        if self.fields:
            print("\nExtracting operation-based fields...")
            for field_name, field_config in self.fields.items():
                try:
                    value = self._dispatch_operation(field_name, field_config)
                    data[field_name] = value
                except Exception as e:
                    print(f"  Warning: Error reading '{field_name}': {e}")
                    data[field_name] = None
        
        else:
            print("\nExtracting standard fields...")
            data.update(self._extract_field_group(self.standard_excel_fields, "standard"))
            print("\nExtracting retail fields...")
            data.update(self._extract_field_group(self.retail_excel_fields, "retail"))
            print("\nExtracting supplier fields...")
            data.update(self._extract_field_group(self.supplier_excel_fields, "supplier"))

        print(f"\nExtracting table data... (found {len(self.table_fields)} tables configured)")
        for table_name, table_config in self.table_fields.items():
            try:
                sheet = table_config["sheet"]
                column_name = table_config["columns"][0]["col"]
                field_name = table_config["columns"][0]["name"]
                count_column = table_config["count_column"]
                limit = table_config.get("limit", None)
                uncategorized_label = table_config.get("uncategorized_label", None)

                print(f"  Reading table '{table_name}' from column '{column_name}'")

                value_counts = excel_utils.count_unique_values(
                    self.excel_path,
                    sheet,
                    column_name,
                    uncategorized_label=uncategorized_label,
                )

                value_counts.sort(key=lambda x: x["count"], reverse=True)

                if uncategorized_label:
                    uncategorized_items = [
                        item for item in value_counts if item["value"] == uncategorized_label
                    ]
                    categorized_items = [
                        item for item in value_counts if item["value"] != uncategorized_label
                    ]
                    value_counts = categorized_items + uncategorized_items

                if limit:
                    value_counts = value_counts[:limit]
                    print(f"  -> Limiting to top {limit} results")

                table_data = []
                for item in value_counts:
                    row = {field_name: item["value"], count_column: item["count"]}
                    table_data.append(row)

                data[table_name] = table_data
                print(f"  -> Found {len(table_data)} unique categories")

            except Exception as e:
                raise ValueError(f"Error reading table '{table_name}': {e}") from e

        print(f"\nExtracted fields: {', '.join([k for k in data if not isinstance(data[k], list)])}")
        return data

    def _dispatch_operation(self, field_name, field_config):
        operation = field_config["operation"]
        source = field_config.get("source", "current")
        excel_path = self.previous_excel_path if source == "previous" else self.excel_path
        
        if source == "previous" and not self.previous_excel_path:
            print(f"  Warning: '{field_name}' requires previous file but none provided")
            return None
        
        sheet = field_config["sheet"]
        filters = field_config.get("filters")
        
        if operation == "count_rows":
            return excel_utils.count_rows(excel_path, sheet, filters)
        
        elif operation == "count_value":
            column = field_config["column"]
            value = field_config["value"]
            return excel_utils.count_column_value(excel_path, sheet, column, value, filters)
        
        elif operation == "sum":
            column = field_config["column"]
            return excel_utils.sum_column(excel_path, sheet, column, filters)
        
        elif operation == "avg_date_diff":
            start_col = field_config["start_column"]
            end_col = field_config["end_column"]
            return excel_utils.calculate_average_resolution_time(
                excel_path, sheet, start_col, end_col, filters
            )
        
        elif operation == "parse_month":
            column = field_config["column"]
            value = excel_utils.read_column_value(excel_path, sheet, column)
            return excel_utils.parse_month_from_date(value)
        
        elif operation == "parse_previous_month":
            column = field_config["column"]
            value = excel_utils.read_column_value(excel_path, sheet, column)
            return excel_utils.parse_previous_month_from_date(value)
        
        elif operation == "read_value":
            column = field_config["column"]
            return excel_utils.read_column_value(excel_path, sheet, column)
        
        else:
            raise ValueError(f"Unknown operation '{operation}' for field '{field_name}'")

    def _extract_field_group(self, field_group, group_name):
        if not field_group:
            return {}

        data = {}

        default_sheet = field_group.get("sheet")

        for field_name, field_config in field_group.items():
            if field_name == "sheet":
                continue

            if not isinstance(field_config, dict):
                continue

            try:
                sheet = field_config.get("sheet", default_sheet)
                cell = field_config.get("cell") 

                is_prev_field = "_prev_" in field_name
                excel_path = (
                    self.previous_excel_path
                    if is_prev_field and self.previous_excel_path
                    else self.excel_path
                )

                if is_prev_field and not self.previous_excel_path:
                    print(
                        f"  Warning: '{field_name}' is a previous month field but no previous Excel file provided"
                    )
                    data[field_name] = None
                    continue

                if field_name.endswith("_req"):
                    value = excel_utils.count_rows(excel_path, sheet)
                    data[field_name] = value
                    continue

                
                if field_name.endswith("_sat_c"):
                    if cell is None or cell == "":
                        print(f"  Skipping '{field_name}': no cell configuration")
                        data[field_name] = None
                        continue
                    print(
                        f"  Reading '{field_name}' from {'previous' if is_prev_field else 'current'} month file, sheet '{sheet}', column '{cell}'"
                    )
                    value = excel_utils.count_column_value(
                        excel_path, sheet, cell, "good with comment"
                    )
                    print(f"  → Found {value} occurrences")
                    data[field_name] = value
                    continue
                    
                elif field_name.endswith("_sat"):
                    if cell is None or cell == "":
                        print(f"  Skipping '{field_name}': no cell configuration")
                        data[field_name] = None
                        continue
                    print(
                        f"  Reading '{field_name}' from {'previous' if is_prev_field else 'current'} month file, sheet '{sheet}', column '{cell}'"
                    )
                    value = excel_utils.count_column_value(
                        excel_path, sheet, cell, "good"
                    )
                    print(f"  → Found {value} occurrences")
                    data[field_name] = value
                    continue

                if field_name.endswith("_reso"):
                    print(
                        f"  Reading '{field_name}' from {'previous' if is_prev_field else 'current'} month file, sheet '{sheet}'"
                    )
                    value = excel_utils.calculate_average_resolution_time(
                        excel_path, sheet
                    )
                    print(f"  → Found {value} days average resolution time")
                    data[field_name] = value
                    continue

                if cell is None or cell == "":
                    print(f"  Skipping '{field_name}': no cell configuration")
                    data[field_name] = None
                    continue

                value = excel_utils.read_column_value(excel_path, sheet, cell)

                if field_name == "month":
                    value = excel_utils.parse_month_from_date(value)
                elif field_name == "prev_month":
                    value = excel_utils.parse_previous_month_from_date(value)

                data[field_name] = value

            except Exception as e:
                print(f"  Warning: Error reading '{field_name}' from {group_name}: {e}")
                data[field_name] = None

        return data
