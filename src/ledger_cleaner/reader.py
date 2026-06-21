import os
import glob
from typing import List, Dict, Tuple, Optional
import pandas as pd
from datetime import datetime
import re

from .config import ValidationConfig
from .models import LedgerRecord, FileParseResult, ValidationIssue


class ExcelReader:
    def __init__(self, config: Optional[ValidationConfig] = None):
        self.config = config or ValidationConfig()

    def find_excel_files(self, folder_path: str) -> List[str]:
        patterns = ["*.xlsx", "*.xls"]
        files = []
        for pattern in patterns:
            files.extend(glob.glob(os.path.join(folder_path, pattern)))
        files = [f for f in files if not os.path.basename(f).startswith("~$")]
        return sorted(files)

    def detect_columns(self, df: pd.DataFrame) -> Dict[str, str]:
        detected = {}
        df_columns = [str(col).strip() for col in df.columns]

        for standard_name, aliases in self.config.column_mappings.items():
            for alias in aliases:
                for col in df_columns:
                    if alias.lower() == col.lower():
                        detected[standard_name] = col
                        break
                if standard_name in detected:
                    break

        return detected

    def parse_date(self, value: Any) -> Optional[datetime]:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None

        if isinstance(value, datetime):
            return value

        if isinstance(value, pd.Timestamp):
            return value.to_pydatetime()

        date_str = str(value).strip()
        if not date_str or date_str.lower() in ["nan", "nat", "none", "null"]:
            return None

        date_str = re.sub(r'[\s\u3000]+', '', date_str)

        for fmt in self.config.date_formats:
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue

        try:
            return pd.to_datetime(date_str).to_pydatetime()
        except (ValueError, TypeError):
            return None

    def parse_amount(self, value: Any) -> Tuple[Optional[float], Optional[str]]:
        if value is None or (isinstance(value, float) and pd.isna(value)):
            return None, "金额为空"

        if isinstance(value, (int, float)):
            if pd.isna(value):
                return None, "金额为空"
            return float(value), None

        amount_str = str(value).strip()
        if not amount_str or amount_str.lower() in ["nan", "nat", "none", "null"]:
            return None, "金额为空"

        original_str = amount_str

        amount_str = amount_str.replace(",", "").replace("，", "")
        amount_str = amount_str.replace("￥", "").replace("¥", "")
        amount_str = amount_str.replace("元", "").replace("RMB", "").replace("rmb", "")
        amount_str = amount_str.replace(" ", "").replace("\u3000", "")

        negative = False
        if amount_str.startswith("-") or amount_str.startswith("("):
            negative = True
            amount_str = amount_str.lstrip("-").lstrip("(").rstrip(")")

        try:
            amount = float(amount_str)
            if negative:
                amount = -amount
            return amount, None
        except ValueError:
            return None, f"金额格式不正确: {original_str}"

    def normalize_status(self, status: str) -> str:
        if not status:
            return ""

        status_lower = status.strip().lower()

        for standard, aliases in self.config.standard_status.items():
            if status_lower in [a.lower() for a in aliases]:
                return standard

        for standard, aliases in self.config.standard_status.items():
            for alias in aliases:
                if alias.lower() in status_lower or status_lower in alias.lower():
                    return standard

        return status.strip()

    def is_standard_status(self, status: str) -> bool:
        return status.strip() in self.config.get_standard_status_set()

    def read_excel_file(self, file_path: str) -> FileParseResult:
        file_name = os.path.basename(file_path)
        result = FileParseResult(file_name=file_name)

        try:
            df = pd.read_excel(file_path)
            result.raw_data = df
        except Exception as e:
            result.parse_success = False
            result.error_message = f"文件读取失败: {str(e)}"
            result.issues.append(ValidationIssue(
                file_name=file_name,
                row_number=0,
                column="",
                issue_type="文件错误",
                description=f"无法读取Excel文件: {str(e)}",
                suggestion="请检查文件是否损坏或格式是否正确",
            ))
            return result

        if df.empty:
            result.issues.append(ValidationIssue(
                file_name=file_name,
                row_number=0,
                column="",
                issue_type="数据为空",
                description="Excel文件中没有数据",
                suggestion="请确认文件中是否包含台账数据",
            ))
            return result

        detected_columns = self.detect_columns(df)
        missing_columns = [col for col in self.config.required_columns if col not in detected_columns]

        if missing_columns:
            result.issues.append(ValidationIssue(
                file_name=file_name,
                row_number=0,
                column="",
                issue_type="缺少列",
                description=f"缺少必要的列: {', '.join(missing_columns)}",
                suggestion=f"请添加以下列或检查列名是否正确: {', '.join(missing_columns)}。可使用的列名别名请参考配置文档。",
            ))
            return result

        for idx, row in df.iterrows():
            row_num = idx + 2

            record = LedgerRecord(
                file_name=file_name,
                row_number=row_num,
                编号="",
            )

            record.原始数据 = {}
            for col in detected_columns:
                record.原始数据[col] = row[detected_columns[col]]

            no_value = row[detected_columns["编号"]]
            if pd.isna(no_value) or str(no_value).strip() == "":
                result.issues.append(ValidationIssue(
                    file_name=file_name,
                    row_number=row_num,
                    column="编号",
                    issue_type="编号缺失",
                    description="变更/洽商编号为空",
                    suggestion="请补充唯一的编号，建议格式: 项目缩写-年份-序号",
                ))
                record.编号 = f"缺失_{row_num}"
            else:
                record.编号 = str(no_value).strip()

            date_value = row[detected_columns["日期"]]
            parsed_date = self.parse_date(date_value)
            if parsed_date is None:
                result.issues.append(ValidationIssue(
                    file_name=file_name,
                    row_number=row_num,
                    column="日期",
                    issue_type="日期缺失或格式错误",
                    description=f"日期值无法解析: {date_value}",
                    suggestion="请补充日期，推荐格式: YYYY-MM-DD (如 2024-01-15)",
                    actual_value=date_value,
                ))
            record.日期 = parsed_date

            name_value = row[detected_columns["名称"]]
            if pd.isna(name_value) or str(name_value).strip() == "":
                result.issues.append(ValidationIssue(
                    file_name=file_name,
                    row_number=row_num,
                    column="名称",
                    issue_type="名称缺失",
                    description="变更/洽商名称为空",
                    suggestion="请补充名称，简要说明变更内容",
                ))
                record.名称 = ""
            else:
                record.名称 = str(name_value).strip()

            amount_value = row[detected_columns["金额"]]
            parsed_amount, amount_error = self.parse_amount(amount_value)
            if amount_error:
                result.issues.append(ValidationIssue(
                    file_name=file_name,
                    row_number=row_num,
                    column="金额",
                    issue_type="金额格式错误",
                    description=amount_error,
                    suggestion="请填写正确的数字金额，可包含千分位和货币符号",
                    actual_value=amount_value,
                ))
            record.金额 = parsed_amount

            status_value = row[detected_columns["状态"]]
            if pd.isna(status_value) or str(status_value).strip() == "":
                result.issues.append(ValidationIssue(
                    file_name=file_name,
                    row_number=row_num,
                    column="状态",
                    issue_type="状态缺失",
                    description="审批状态为空",
                    suggestion=f"请从以下状态中选择: {', '.join(self.config.get_standard_status_set())}",
                ))
                record.状态 = ""
                record.标准状态 = ""
            else:
                record.状态 = str(status_value).strip()
                normalized = self.normalize_status(record.状态)
                record.标准状态 = normalized

                if not self.is_standard_status(normalized):
                    result.issues.append(ValidationIssue(
                        file_name=file_name,
                        row_number=row_num,
                        column="状态",
                        issue_type="状态不规范",
                        description=f"状态写法不规范: {record.状态}",
                        suggestion=f"请调整为标准状态之一: {', '.join(self.config.get_standard_status_set())}",
                        actual_value=record.状态,
                    ))

            if "附件" in detected_columns:
                attach_value = row[detected_columns["附件"]]
                record.附件 = "" if pd.isna(attach_value) else str(attach_value).strip()

            if "项目名称" in detected_columns:
                proj_value = row[detected_columns["项目名称"]]
                record.项目名称 = "" if pd.isna(proj_value) else str(proj_value).strip()

            result.records.append(record)

        return result

    def read_folder(self, folder_path: str) -> List[FileParseResult]:
        results = []
        excel_files = self.find_excel_files(folder_path)

        if not excel_files:
            return results

        for file_path in excel_files:
            result = self.read_excel_file(file_path)
            results.append(result)

        return results
