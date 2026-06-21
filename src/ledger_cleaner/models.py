from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional, Dict, Any, List as _List
from datetime import datetime
import pandas as pd


@dataclass
class ValidationIssue:
    file_name: str
    row_number: int
    column: str
    issue_type: str
    description: str
    suggestion: str
    actual_value: Any = None

    def to_dict(self) -> Dict[str, Any]:
        return {
            "文件名称": self.file_name,
            "行号": self.row_number,
            "列名": self.column,
            "问题类型": self.issue_type,
            "问题描述": self.description,
            "建议补充": self.suggestion,
            "当前值": str(self.actual_value) if self.actual_value is not None else "",
        }


@dataclass
class LedgerRecord:
    file_name: str
    row_number: int
    编号: str
    日期: Optional[datetime] = None
    名称: str = ""
    金额: Optional[float] = None
    状态: str = ""
    标准状态: str = ""
    附件: str = ""
    项目名称: str = ""
    原始数据: Dict[str, Any] = field(default_factory=dict)

    def to_dict(self) -> Dict[str, Any]:
        return {
            "文件名称": self.file_name,
            "行号": self.row_number,
            "编号": self.编号,
            "日期": self.日期.strftime("%Y-%m-%d") if self.日期 else "",
            "名称": self.名称,
            "金额": self.金额 if self.金额 is not None else 0,
            "状态": self.状态,
            "标准状态": self.标准状态,
            "附件": self.附件,
            "项目名称": self.项目名称,
        }


@dataclass
class FileParseResult:
    file_name: str
    records: list = field(default_factory=list)
    issues: list = field(default_factory=list)
    raw_data: Optional[pd.DataFrame] = None
    parse_success: bool = True
    error_message: str = ""
