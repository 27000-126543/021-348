from __future__ import annotations

from dataclasses import dataclass, field
from typing import Dict, List, Set, Optional as _Optional

STANDARD_STATUS = {
    "未报审": ["未报审", "未提交", "待报审", "待提交", "未上报", "草稿"],
    "已报审": ["已报审", "已提交", "已上报", "报审中", "审核中", "已申报"],
    "已确认": ["已确认", "已批复", "已审核", "已批准", "已同意", "确认", "批复"],
    "已进结算": ["已进结算", "已结算", "结算中", "已计入", "已纳入"],
}

COLUMN_MAPPINGS = {
    "编号": ["编号", "变更编号", "洽商编号", "单据编号", "序号", "no", "NO", "No."],
    "日期": ["日期", "变更日期", "洽商日期", "发生日期", "填报日期", "创建日期", "date", "Date"],
    "名称": ["名称", "变更名称", "洽商名称", "内容摘要", "摘要", "事项", "item", "Item"],
    "金额": ["金额", "变更金额", "洽商金额", "造价", "费用", "金额(元)", "amount", "Amount", "价格"],
    "状态": ["状态", "审批状态", "审核状态", "流程状态", "status", "Status"],
    "附件": ["附件", "审批附件", "相关附件", "佐证材料", "证明材料", "attachment"],
    "项目名称": ["项目名称", "项目", "工程名称", "project", "Project"],
}

REQUIRED_COLUMNS = ["编号", "日期", "名称", "金额", "状态"]

DATE_FORMATS = [
    "%Y-%m-%d", "%Y/%m/%d", "%Y.%m.%d",
    "%Y年%m月%d日",
    "%Y-%m", "%Y/%m", "%Y.%m",
    "%Y年%m月",
    "%m/%d/%Y", "%d/%m/%Y",
    "%Y%m%d",
]

AMOUNT_THRESHOLD = 50000.0


@dataclass
class ValidationConfig:
    amount_threshold: float = AMOUNT_THRESHOLD
    standard_status: Dict[str, List[str]] = field(default_factory=lambda: STANDARD_STATUS)
    required_columns: List[str] = field(default_factory=lambda: REQUIRED_COLUMNS)
    column_mappings: Dict[str, List[str]] = field(default_factory=lambda: COLUMN_MAPPINGS)
    date_formats: List[str] = field(default_factory=lambda: DATE_FORMATS)

    def get_standard_status_set(self) -> Set[str]:
        return set(self.standard_status.keys())

    def get_all_status_aliases(self) -> Set[str]:
        aliases = set()
        for values in self.standard_status.values():
            aliases.update(v.lower() for v in values)
        return aliases
