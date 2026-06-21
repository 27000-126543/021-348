from __future__ import annotations

import os
from typing import List, Dict, Optional
import pandas as pd
from datetime import datetime

from .models import ValidationIssue, LedgerRecord, FileParseResult
from .config import ValidationConfig


class ReportGenerator:
    def __init__(self, config: ValidationConfig = None):
        self.config = config or ValidationConfig()

    def generate_issue_report(self, issues: List[ValidationIssue],
                              project_name: str,
                              project_month: str,
                              output_path: str) -> str:
        if not issues:
            return ""

        try:
            os.makedirs(output_path, exist_ok=True)
        except Exception as e:
            raise RuntimeError(f"无法创建输出目录: {output_path}, 错误: {str(e)}")

        try:
            df = pd.DataFrame([issue.to_dict() for issue in issues])
        except Exception as e:
            raise RuntimeError(f"生成问题清单数据时出错: {str(e)}")

        column_order = ["文件名称", "行号", "列名", "问题类型", "问题描述", "建议补充", "当前值"]
        df = df.reindex(columns=column_order)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_project_name = "".join(c for c in project_name if c.isalnum() or c in ("-", "_"))
        file_name = f"{safe_project_name}_{project_month}_问题清单_{timestamp}.xlsx"
        full_path = os.path.join(output_path, file_name)

        try:
            with pd.ExcelWriter(full_path, engine='openpyxl') as writer:
                df.to_excel(writer, sheet_name="问题清单", index=False)

                try:
                    worksheet = writer.sheets["问题清单"]
                    for column in worksheet.columns:
                        max_length = 0
                        column_letter = column[0].column_letter
                        for cell in column:
                            try:
                                if len(str(cell.value)) > max_length:
                                    max_length = len(str(cell.value))
                            except:
                                pass
                        adjusted_width = min(max_length + 2, 60)
                        worksheet.column_dimensions[column_letter].width = adjusted_width
                except Exception:
                    pass

                try:
                    summary_df = self._get_issue_summary_dataframe(issues)
                    summary_df.to_excel(writer, sheet_name="问题汇总", index=False)
                except Exception:
                    pass
        except PermissionError:
            raise RuntimeError(f"无法写入文件 {full_path}，文件可能被Excel打开。请关闭Excel后重试。")
        except Exception as e:
            raise RuntimeError(f"写入Excel文件时出错: {str(e)}")

        return full_path

    def _get_issue_summary_dataframe(self, issues: List[ValidationIssue]) -> pd.DataFrame:
        summary: Dict[str, int] = {}
        file_summary: Dict[str, Dict[str, int]] = {}

        for issue in issues:
            summary[issue.issue_type] = summary.get(issue.issue_type, 0) + 1

            if issue.file_name not in file_summary:
                file_summary[issue.file_name] = {}
            file_summary[issue.file_name][issue.issue_type] = \
                file_summary[issue.file_name].get(issue.issue_type, 0) + 1

        rows = []
        rows.append({"项目": "问题类型汇总", "数量": ""})
        for issue_type, count in sorted(summary.items()):
            rows.append({"项目": f"  - {issue_type}", "数量": count})
        rows.append({"项目": "合计问题数", "数量": len(issues)})
        rows.append({"项目": "", "数量": ""})
        rows.append({"项目": "按文件汇总", "数量": ""})

        for file_name, file_issues in sorted(file_summary.items()):
            file_total = sum(file_issues.values())
            rows.append({"项目": f"  {file_name}", "数量": file_total})
            for issue_type, count in sorted(file_issues.items()):
                rows.append({"项目": f"    - {issue_type}", "数量": count})

        return pd.DataFrame(rows)

    def print_issue_summary(self, issues: List[ValidationIssue]) -> None:
        if not issues:
            print("\n✅ 恭喜！未发现数据质量问题。")
            return

        print(f"\n{'='*60}")
        print(f"📋 数据质量检查报告 - 共发现 {len(issues)} 个问题")
        print(f"{'='*60}")

        from collections import defaultdict
        type_summary = defaultdict(int)
        file_summary = defaultdict(int)

        for issue in issues:
            type_summary[issue.issue_type] += 1
            file_summary[issue.file_name] += 1

        print("\n📊 按问题类型汇总:")
        for issue_type, count in sorted(type_summary.items()):
            print(f"   {issue_type}: {count} 个")

        print("\n📁 按文件汇总:")
        for file_name, count in sorted(file_summary.items()):
            print(f"   {file_name}: {count} 个")

        print(f"\n{'='*60}")
        print("📝 问题明细 (前10条):")
        print(f"{'='*60}")

        for i, issue in enumerate(issues[:10], 1):
            print(f"\n{i}. [{issue.file_name}] 第{issue.row_number}行 - {issue.issue_type}")
            print(f"   列名: {issue.column}")
            print(f"   描述: {issue.description}")
            print(f"   建议: {issue.suggestion}")
            if issue.actual_value:
                print(f"   当前值: {issue.actual_value}")

        if len(issues) > 10:
            print(f"\n... 还有 {len(issues) - 10} 条问题请查看完整报告文件。")

        print(f"\n{'='*60}\n")
