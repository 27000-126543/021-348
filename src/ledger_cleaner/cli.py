from __future__ import annotations

import os
import sys
import traceback
import functools
import click
from typing import Optional

from .config import ValidationConfig
from .reader import ExcelReader
from .validator import DataValidator
from .report import ReportGenerator
from .summarizer import LedgerSummarizer


def handle_errors(func):
    @functools.wraps(func)
    def wrapper(*args, **kwargs):
        try:
            return func(*args, **kwargs)
        except click.exceptions.ClickException:
            raise
        except SystemExit:
            raise
        except KeyboardInterrupt:
            click.echo()
            click.echo("⏹️  操作已取消")
            sys.exit(130)
        except Exception as e:
            click.echo()
            click.echo("=" * 60, err=True)
            click.echo(f"❌ 运行出错: {str(e)}", err=True)
            click.echo("=" * 60, err=True)
            click.echo()
            click.echo("ℹ️  详细错误信息:", err=True)
            click.echo("-" * 60, err=True)
            traceback.print_exc()
            click.echo("-" * 60, err=True)
            click.echo()
            click.echo("💡 如果此问题持续出现，请检查:", err=True)
            click.echo("   - Excel文件是否正常打开", err=True)
            click.echo("   - 数据格式是否正确", err=True)
            click.echo("   - 是否有足够的磁盘空间", err=True)
            click.echo()
            sys.exit(1)
    return wrapper


@click.group()
@click.version_option(version="1.0.0", prog_name="ledger-cleaner")
@handle_errors
def cli():
    """变更洽商台账整理工具 - 批量清理历史台账数据

    适用于造价员和商务助理在月底对账前快速体检历史数据。

    典型流程:
      1. 将多个项目导出的Excel表格放入同一文件夹
      2. 运行 'ledger check' 检查数据质量问题
      3. 根据问题清单修正原始数据
      4. 重新检查无误后，运行 'ledger export' 生成标准台账
    """
    pass


@cli.command()
@click.option('--input', '-i', required=True, type=click.Path(exists=True, file_okay=False),
              help='包含Excel台账文件的文件夹路径')
@click.option('--project', '-p', required=True, help='项目名称（用于输出文件命名）')
@click.option('--month', '-m', required=True, help='月份，格式: YYYY-MM (如: 2024-01)')
@click.option('--output', '-o', default=None, type=click.Path(file_okay=False),
              help='问题报告输出文件夹，默认为输入文件夹')
@click.option('--threshold', '-t', default=50000, type=float,
              help='大额审批附件阈值（元），默认50000元')
@click.option('--yes', '-y', is_flag=True, help='跳过确认直接输出报告')
@handle_errors
def check(input, project, month, output, threshold, yes):
    """检查台账数据质量，生成问题清单

    检查内容包括：
    - 编号重复（跨文件检查）
    - 日期缺失或格式错误
    - 金额格式不一致
    - 状态写法不规范
    - 大额事项缺少审批附件
    - 同一编号在不同表中状态不一致
    """
    click.echo()
    click.echo("=" * 60)
    click.echo("🔍 变更洽商台账数据质量检查")
    click.echo("=" * 60)
    click.echo()
    click.echo(f"📁 输入文件夹: {input}")
    click.echo(f"🏗️  项目名称: {project}")
    click.echo(f"📅 月份: {month}")
    click.echo(f"💰 大额阈值: {threshold:,.0f} 元")
    click.echo()

    config = ValidationConfig(amount_threshold=threshold)
    reader = ExcelReader(config)
    validator = DataValidator(config)
    reporter = ReportGenerator(config)

    excel_files = reader.find_excel_files(input)
    if not excel_files:
        click.echo("❌ 未找到Excel文件！请检查文件夹路径是否正确。")
        click.echo(f"   支持格式: .xlsx, .xls")
        sys.exit(1)

    click.echo(f"📄 发现 {len(excel_files)} 个Excel文件:")
    for f in excel_files:
        click.echo(f"   - {os.path.basename(f)}")
    click.echo()

    with click.progressbar(length=len(excel_files), label='🔄 正在读取文件') as bar:
        parse_results = []
        for file_path in excel_files:
            result = reader.read_excel_file(file_path)
            parse_results.append(result)
            bar.update(1)

    click.echo()
    click.echo("🔎 正在进行数据校验...")
    issues = validator.validate_all(parse_results, project_month=month)

    reporter.print_issue_summary(issues)

    if not issues:
        click.echo("🎉 数据质量良好，可以直接生成台账汇总！")
        click.echo("   运行: ledger export 生成标准台账文件")
        return

    output_path = output or input

    if not yes:
        click.echo()
        confirm = click.confirm(f"是否生成问题清单Excel报告到 {output_path}？", default=True)
        if not confirm:
            click.echo("已取消生成报告。")
            return

    report_file = reporter.generate_issue_report(issues, project, month, output_path)
    click.echo()
    click.echo(f"✅ 问题清单已生成: {report_file}")
    click.echo()
    click.echo("📋 下一步操作:")
    click.echo("   1. 根据问题清单修正原始Excel文件")
    click.echo("   2. 再次运行 'ledger check' 确认问题已解决")
    click.echo("   3. 确认无误后运行 'ledger export' 生成标准台账")


@cli.command()
@click.option('--input', '-i', required=True, type=click.Path(exists=True, file_okay=False),
              help='包含Excel台账文件的文件夹路径')
@click.option('--project', '-p', required=True, help='项目名称（用于输出文件命名）')
@click.option('--month', '-m', required=True, help='月份，格式: YYYY-MM (如: 2024-01)')
@click.option('--output', '-o', default=None, type=click.Path(file_okay=False),
              help='汇总文件输出文件夹，默认为输入文件夹')
@click.option('--threshold', '-t', default=50000, type=float,
              help='大额审批附件阈值（元），默认50000元')
@click.option('--force', '-f', is_flag=True, help='即使有问题也强制生成汇总（不推荐）')
@handle_errors
def export(input, project, month, output, threshold, force):
    """生成标准台账汇总文件

    按状态分组：未报审、已报审、已确认、已进结算
    额外列出大额无附件预警清单
    """
    click.echo()
    click.echo("=" * 60)
    click.echo("📊 生成标准台账汇总")
    click.echo("=" * 60)
    click.echo()
    click.echo(f"📁 输入文件夹: {input}")
    click.echo(f"🏗️  项目名称: {project}")
    click.echo(f"📅 月份: {month}")
    click.echo(f"💰 大额阈值: {threshold:,.0f} 元")
    click.echo()

    config = ValidationConfig(amount_threshold=threshold)
    reader = ExcelReader(config)
    validator = DataValidator(config)
    summarizer = LedgerSummarizer(config)

    excel_files = reader.find_excel_files(input)
    if not excel_files:
        click.echo("❌ 未找到Excel文件！请检查文件夹路径是否正确。")
        sys.exit(1)

    click.echo(f"📄 读取 {len(excel_files)} 个Excel文件...")

    parse_results = []
    with click.progressbar(excel_files, label='🔄 正在处理') as bar:
        for file_path in bar:
            result = reader.read_excel_file(file_path)
            parse_results.append(result)

    issues = validator.validate_all(parse_results, project_month=month)

    if issues and not force:
        click.echo()
        click.echo(f"⚠️  发现 {len(issues)} 个数据质量问题！")
        click.echo("   建议先运行 'ledger check' 查看并修复问题。")
        click.echo()
        confirm = click.confirm("是否仍要生成汇总文件？（数据可能不准确）", default=False)
        if not confirm:
            click.echo("已取消。")
            return
    elif issues:
        click.echo(f"⚠️  强制模式：忽略 {len(issues)} 个问题继续生成。")

    output_path = output or input

    click.echo()
    click.echo("📈 正在生成汇总...")
    summary_file = summarizer.generate_summary_excel(
        parse_results, project, month, output_path
    )

    click.echo()
    summarizer.print_summary(parse_results, project, month)

    click.echo(f"✅ 台账汇总已生成: {summary_file}")
    click.echo()
    click.echo("📂 汇总文件包含以下工作表:")
    click.echo("   - 统计汇总: 各状态条数和金额合计")
    click.echo("   - 全部数据: 所有台账记录明细")
    click.echo("   - 【未报审】: 待提交的变更洽商")
    click.echo("   - 【已报审】: 已提交待审批的事项")
    click.echo("   - 【已确认】: 审批完成的事项")
    click.echo("   - 【已进结算】: 已计入结算的事项")
    if config.amount_threshold > 0:
        click.echo("   - ⚠️大额无附件: 需补充审批附件的预警清单")
    click.echo()


@cli.command()
@click.option('--path', '-p', default='./input', type=click.Path(file_okay=False),
              help='示例文件生成路径')
@click.option('--include-errors', '-e', is_flag=True,
              help='生成包含常见错误的示例文件（用于测试）')
@handle_errors
def demo(path, include_errors):
    """生成示例台账文件

    创建用于测试的示例Excel文件，可选择包含常见数据错误。
    """
    try:
        import pandas as pd
    except ImportError:
        click.echo("❌ 需要安装 pandas: pip install pandas openpyxl")
        sys.exit(1)

    os.makedirs(path, exist_ok=True)

    demo_data_1 = [
        ["编号", "日期", "名称", "金额", "状态", "附件"],
        ["BQ-2024-001", "2024-01-15", "地下室墙体变更", 125000, "已确认", "审批单001.pdf"],
        ["BQ-2024-002", "2024/01/18", "屋面防水做法调整", 45000, "已报审", "图纸会审记录.pdf"],
        ["BQ-2024-003", "2024.01.20", "电梯井尺寸变更", 80000, "未报审", ""],
        ["BQ-2024-004", "2024年01月22日", "大堂装修方案调整", 156000, "已进结算", "签证单004.pdf"],
    ]

    demo_data_2 = [
        ["变更编号", "变更日期", "变更名称", "造价", "审批状态", "相关附件"],
        ["BQ-2024-005", "2024-01-08", "给排水管线走向调整", 32000, "审核中", ""],
        ["BQ-2024-006", "2024-01-25", "消防喷淋点位增加", 68000, "已批准", "审批单006.pdf"],
        ["BQ-2024-002", "2024-01-18", "屋面防水做法调整", 45000, "已确认", "图纸会审记录.pdf"],
    ]

    if include_errors:
        demo_data_1.extend([
            ["", "2024-01-25", "空调系统调整", 78000, "待审核", ""],
            ["BQ-2024-007", "", "强弱电井扩容", 95000, "已提交", ""],
            ["BQ-2024-008", "2024-01-10", "室外管网调整", "壹拾贰万", "Done", "附件齐全"],
            ["BQ-2024-009", "2024-01-12", "停车场划线方案", 65000, "完成", "有"],
            ["BQ-2024-010", "2024-01-30", "景观绿化变更", 120000, "已报审", ""],
        ])
        demo_data_2.extend([
            ["BQ-2024-011", "2024-02-01", "门禁系统升级", 45000, "未提交", ""],
            ["BQ-2024-012", "2024-01-15", "会议室改造", "85,000元", "已审核", "扫描件.pdf"],
        ])

    df1 = pd.DataFrame(demo_data_1[1:], columns=demo_data_1[0])
    file1 = os.path.join(path, "项目A_台账1.xlsx")
    df1.to_excel(file1, index=False)

    df2 = pd.DataFrame(demo_data_2[1:], columns=demo_data_2[0])
    file2 = os.path.join(path, "项目A_台账2.xlsx")
    df2.to_excel(file2, index=False)

    click.echo()
    click.echo("=" * 60)
    click.echo("✅ 示例文件已生成")
    click.echo("=" * 60)
    click.echo()
    click.echo(f"📁 生成路径: {os.path.abspath(path)}")
    click.echo()
    click.echo("📄 生成的文件:")
    click.echo(f"   - {os.path.basename(file1)} ({len(df1)} 条记录)")
    click.echo(f"   - {os.path.basename(file2)} ({len(df2)} 条记录)")
    click.echo()
    if include_errors:
        click.echo("⚠️  已包含常见数据错误，用于测试检查功能:")
        click.echo("   - 编号缺失、日期缺失")
        click.echo("   - 金额格式不统一（中文大写、带单位）")
        click.echo("   - 状态写法混乱（英文、口语化）")
        click.echo("   - 编号跨文件重复")
        click.echo("   - 大额无附件")
        click.echo("   - 日期超出月份范围")
        click.echo()
    click.echo("🚀 现在可以运行:")
    click.echo(f"   ledger check -i {path} -p 项目A -m 2024-01")
    click.echo()


@cli.command()
@handle_errors
def status():
    """查看标准状态列表和列名映射"""
    config = ValidationConfig()

    click.echo()
    click.echo("=" * 60)
    click.echo("📋 标准配置参考")
    click.echo("=" * 60)

    click.echo()
    click.echo("🔄 标准状态及识别别名:")
    for standard, aliases in config.standard_status.items():
        click.echo(f"   ✅ {standard}: {', '.join(aliases)}")

    click.echo()
    click.echo("📝 支持的列名（别名自动识别）:")
    for standard, aliases in config.column_mappings.items():
        required = "*" if standard in config.required_columns else " "
        click.echo(f"   {required} {standard}: {', '.join(aliases[:5])}")
        if len(aliases) > 5:
            click.echo(f"            等 {len(aliases)} 种写法")

    click.echo()
    click.echo(f"💰 默认大额阈值: {config.amount_threshold:,.0f} 元")
    click.echo("   可通过 --threshold 参数调整")
    click.echo()

    click.echo("📅 支持的日期格式:")
    for fmt in config.date_formats[:8]:
        sample = ""
        if fmt == "%Y-%m-%d":
            sample = " (如 2024-01-15)"
        elif fmt == "%Y年%m月%d日":
            sample = " (如 2024年01月15日)"
        elif fmt == "%Y%m%d":
            sample = " (如 20240115)"
        click.echo(f"   {fmt}{sample}")
    click.echo("   以及其他 pandas 可解析的格式")
    click.echo()


if __name__ == "__main__":
    cli()
