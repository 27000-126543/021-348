from __future__ import annotations

from typing import List, Dict, Tuple, Optional
from collections import defaultdict
from datetime import datetime

from .config import ValidationConfig
from .models import ValidationIssue, LedgerRecord, FileParseResult


class DataValidator:
    def __init__(self, config: Optional[ValidationConfig] = None):
        self.config = config or ValidationConfig()

    def check_duplicate_numbers(self, parse_results: List[FileParseResult]) -> List[ValidationIssue]:
        issues = []
        number_locations: Dict[str, List[Tuple[str, int, str]]] = defaultdict(list)

        for result in parse_results:
            for record in result.records:
                if record.编号 and not record.编号.startswith("缺失_"):
                    number_locations[record.编号].append(
                        (result.file_name, record.row_number, record.名称)
                    )

        for no, locations in number_locations.items():
            if len(locations) > 1:
                for idx, (file_name, row_num, name) in enumerate(locations):
                    other_files = ", ".join(
                        [f"{loc[0]}第{loc[1]}行" for i, loc in enumerate(locations) if i != idx]
                    )
                    issues.append(ValidationIssue(
                        file_name=file_name,
                        row_number=row_num,
                        column="编号",
                        issue_type="编号重复",
                        description=f"编号 '{no}' 在以下位置重复出现: {other_files}",
                        suggestion="请核对编号是否为同一事项。如确为同一事项，建议合并；如为不同事项，请修改编号确保唯一。",
                        actual_value=no,
                    ))

        return issues

    def check_date_range(self, parse_results: List[FileParseResult],
                         project_month: Optional[str] = None) -> List[ValidationIssue]:
        issues = []

        if not project_month:
            return issues

        try:
            month_start = datetime.strptime(project_month, "%Y-%m")
            next_month = month_start.month + 1
            next_year = month_start.year
            if next_month > 12:
                next_month = 1
                next_year += 1
            month_end = datetime(next_year, next_month, 1)
        except ValueError:
            return issues

        for result in parse_results:
            for record in result.records:
                if record.日期:
                    if record.日期 < month_start or record.日期 >= month_end:
                        issues.append(ValidationIssue(
                            file_name=result.file_name,
                            row_number=record.row_number,
                            column="日期",
                            issue_type="日期超出范围",
                            description=f"日期 {record.日期.strftime('%Y-%m-%d')} 不在 {project_month} 月份范围内",
                            suggestion="请核对日期是否正确。如为其他月份事项，建议移至对应月份台账。",
                            actual_value=record.日期.strftime("%Y-%m-%d"),
                        ))

        return issues

    def check_amount_threshold_attachment(self, parse_results: List[FileParseResult]) -> List[ValidationIssue]:
        issues = []
        threshold = self.config.amount_threshold

        for result in parse_results:
            for record in result.records:
                if record.金额 is not None and abs(record.金额) >= threshold:
                    if not record.附件 or record.附件.lower() in ["无", "没有", "缺", "未提供", ""]:
                        issues.append(ValidationIssue(
                            file_name=result.file_name,
                            row_number=record.row_number,
                            column="附件",
                            issue_type="大额无附件",
                            description=f"金额 {record.金额:,.2f} 元超过阈值 {threshold:,.2f} 元，但缺少审批附件",
                            suggestion="请补充审批扫描件、会议纪要、签证单等佐证材料，标记附件名称或编号。",
                            actual_value=f"{record.金额:,.2f} 元",
                        ))

        return issues

    def check_status_consistency(self, parse_results: List[FileParseResult]) -> List[ValidationIssue]:
        issues = []
        status_records: Dict[str, List[Tuple[str, int, str, str]]] = defaultdict(list)

        for result in parse_results:
            for record in result.records:
                if record.编号 and not record.编号.startswith("缺失_") and record.标准状态:
                    status_records[record.编号].append(
                        (result.file_name, record.row_number, record.状态, record.标准状态)
                    )

        for no, records in status_records.items():
            if len(records) > 1:
                statuses = set(r[3] for r in records if r[3])
                if len(statuses) > 1:
                    for idx, (file_name, row_num, original_status, std_status) in enumerate(records):
                        other_statuses = ", ".join(
                            [f"{r[0]}第{r[1]}行({r[2]})" for i, r in enumerate(records) if i != idx]
                        )
                        issues.append(ValidationIssue(
                            file_name=file_name,
                            row_number=row_num,
                            column="状态",
                            issue_type="状态不一致",
                            description=f"同一编号 '{no}' 在不同表中状态不一致。本表状态: {original_status}，其他表: {other_statuses}",
                            suggestion="请核对审批实际进度，统一各表中的状态。",
                            actual_value=original_status,
                        ))

        return issues

    def validate_all(self, parse_results: List[FileParseResult],
                     project_month: Optional[str] = None) -> List[ValidationIssue]:
        all_issues = []

        try:
            for result in parse_results:
                all_issues.extend(result.issues)
        except Exception as e:
            all_issues.append(ValidationIssue(
                file_name="系统",
                row_number=0,
                column="",
                issue_type="校验错误",
                description=f"收集已有问题时出错: {str(e)}",
                suggestion="请检查文件数据格式是否正常。",
            ))

        try:
            all_issues.extend(self.check_duplicate_numbers(parse_results))
        except Exception as e:
            all_issues.append(ValidationIssue(
                file_name="系统",
                row_number=0,
                column="编号",
                issue_type="校验错误",
                description=f"检查编号重复时出错: {str(e)}",
                suggestion="请检查编号列的数据是否正常。",
            ))

        try:
            all_issues.extend(self.check_date_range(parse_results, project_month))
        except Exception as e:
            all_issues.append(ValidationIssue(
                file_name="系统",
                row_number=0,
                column="日期",
                issue_type="校验错误",
                description=f"检查日期范围时出错: {str(e)}",
                suggestion="请检查日期列的数据格式是否正常。",
            ))

        try:
            all_issues.extend(self.check_amount_threshold_attachment(parse_results))
        except Exception as e:
            all_issues.append(ValidationIssue(
                file_name="系统",
                row_number=0,
                column="金额/附件",
                issue_type="校验错误",
                description=f"检查大额附件时出错: {str(e)}",
                suggestion="请检查金额和附件列的数据是否正常。",
            ))

        try:
            all_issues.extend(self.check_status_consistency(parse_results))
        except Exception as e:
            all_issues.append(ValidationIssue(
                file_name="系统",
                row_number=0,
                column="状态",
                issue_type="校验错误",
                description=f"检查状态一致性时出错: {str(e)}",
                suggestion="请检查状态列的数据是否正常。",
            ))

        try:
            all_issues.sort(key=lambda x: (x.file_name, x.row_number))
        except Exception:
            pass

        return all_issues

    def get_issue_summary(self, issues: List[ValidationIssue]) -> Dict[str, int]:
        summary: Dict[str, int] = defaultdict(int)
        for issue in issues:
            summary[issue.issue_type] += 1
        return dict(summary)
