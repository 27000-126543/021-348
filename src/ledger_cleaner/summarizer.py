import os
from typing import List, Dict, Optional
import pandas as pd
from datetime import datetime
from collections import defaultdict

from .models import LedgerRecord, FileParseResult
from .config import ValidationConfig


class LedgerSummarizer:
    def __init__(self, config: Optional[ValidationConfig] = None):
        self.config = config or ValidationConfig()

    def collect_all_records(self, parse_results: List[FileParseResult],
                            project_name: str) -> List[LedgerRecord]:
        records = []
        for result in parse_results:
            for record in result.records:
                if not record.项目名称:
                    record.项目名称 = project_name
                records.append(record)
        return records

    def group_by_status(self, records: List[LedgerRecord]) -> Dict[str, List[LedgerRecord]]:
        groups: Dict[str, List[LedgerRecord]] = defaultdict(list)
        for status in self.config.get_standard_status_set():
            groups[status] = []
        groups["其他"] = []

        for record in records:
            if record.标准状态 in groups:
                groups[record.标准状态].append(record)
            else:
                groups["其他"].append(record)

        return groups

    def find_high_amount_no_attachment(self, records: List[LedgerRecord]) -> List[LedgerRecord]:
        threshold = self.config.amount_threshold
        result = []
        for record in records:
            if record.金额 is not None and abs(record.金额) >= threshold:
                if not record.附件 or record.附件.lower() in ["无", "没有", "缺", "未提供", ""]:
                    result.append(record)
        return result

    def calculate_statistics(self, groups: Dict[str, List[LedgerRecord]]) -> pd.DataFrame:
        rows = []
        for status, records in groups.items():
            count = len(records)
            total_amount = sum(r.金额 for r in records if r.金额 is not None)
            rows.append({
                "状态": status,
                "条数": count,
                "金额合计(元)": total_amount,
            })

        total_count = sum(r["条数"] for r in rows)
        total_amount = sum(r["金额合计(元)"] for r in rows)
        rows.append({
            "状态": "合计",
            "条数": total_count,
            "金额合计(元)": total_amount,
        })

        return pd.DataFrame(rows)

    def _records_to_dataframe(self, records: List[LedgerRecord]) -> pd.DataFrame:
        data = []
        for record in records:
            row = {
                "项目名称": record.项目名称,
                "文件来源": record.file_name,
                "原始行号": record.row_number,
                "编号": record.编号,
                "日期": record.日期.strftime("%Y-%m-%d") if record.日期 else "",
                "名称": record.名称,
                "金额(元)": record.金额 if record.金额 is not None else 0,
                "原始状态": record.状态,
                "标准状态": record.标准状态,
                "附件": record.附件,
            }
            data.append(row)
        return pd.DataFrame(data)

    def generate_summary_excel(self, parse_results: List[FileParseResult],
                               project_name: str,
                               project_month: str,
                               output_path: str) -> str:
        all_records = self.collect_all_records(parse_results, project_name)
        groups = self.group_by_status(all_records)
        high_amount_no_attach = self.find_high_amount_no_attachment(all_records)
        statistics_df = self.calculate_statistics(groups)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        file_name = f"{project_name}_{project_month}_台账汇总_{timestamp}.xlsx"
        full_path = os.path.join(output_path, file_name)

        with pd.ExcelWriter(full_path, engine='openpyxl') as writer:
            statistics_df.to_excel(writer, sheet_name="统计汇总", index=False)
            self._auto_adjust_column_width(writer.sheets["统计汇总"])

            all_df = self._records_to_dataframe(all_records)
            if not all_df.empty:
                all_df = all_df.sort_values(by=["标准状态", "日期"], na_position='last')
            all_df.to_excel(writer, sheet_name="全部数据", index=False)
            self._auto_adjust_column_width(writer.sheets["全部数据"])

            for status, records in groups.items():
                if records:
                    df = self._records_to_dataframe(records)
                    df.to_excel(writer, sheet_name=f"【{status}】", index=False)
                    self._auto_adjust_column_width(writer.sheets[f"【{status}】"])

            if high_amount_no_attach:
                warning_df = self._records_to_dataframe(high_amount_no_attach)
                warning_df["警告"] = f"金额超过{self.config.amount_threshold:,.0f}元但缺少附件"
                warning_df.to_excel(writer, sheet_name="⚠️大额无附件", index=False)
                self._auto_adjust_column_width(writer.sheets["⚠️大额无附件"])

        return full_path

    def _auto_adjust_column_width(self, worksheet) -> None:
        for column in worksheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column_letter].width = adjusted_width

    def print_summary(self, parse_results: List[FileParseResult],
                      project_name: str,
                      project_month: str) -> None:
        all_records = self.collect_all_records(parse_results, project_name)
        groups = self.group_by_status(all_records)
        high_amount_no_attach = self.find_high_amount_no_attachment(all_records)
        statistics_df = self.calculate_statistics(groups)

        print(f"\n{'='*60}")
        print(f"📊 台账汇总 - {project_name} {project_month}")
        print(f"{'='*60}")

        print("\n📈 状态统计:")
        for _, row in statistics_df.iterrows():
            if row["状态"] == "合计":
                print(f"   {'─'*40}")
            status_marker = "📌" if row["状态"] == "未报审" else \
                            "📤" if row["状态"] == "已报审" else \
                            "✅" if row["状态"] == "已确认" else \
                            "💰" if row["状态"] == "已进结算" else \
                            "📋" if row["状态"] == "合计" else "❓"
            print(f"   {status_marker} {row['状态']}: {row['条数']} 条, 金额 {row['金额合计(元)']:,.2f} 元")

        if high_amount_no_attach:
            print(f"\n⚠️  预警：发现 {len(high_amount_no_attach)} 条大额无附件事项")
            print(f"   （金额超过 {self.config.amount_threshold:,.0f} 元但缺少审批附件）")
            for record in high_amount_no_attach:
                print(f"   - [{record.file_name}] {record.编号} {record.名称}: {record.金额:,.2f} 元")

        print(f"\n{'='*60}\n")
