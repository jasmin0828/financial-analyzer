"""端到端测试 - 用真实财务数据 + 年初余额自动补充"""
from analyzer import analyze_files, generate_report
from pathlib import Path

files = [
    '/Users/jasmin0828/Downloads/蓉惠芯创基础资料/9、2024-2026年3月财务报表/2025年资产负债表.xlsx',
    '/Users/jasmin0828/Downloads/蓉惠芯创基础资料/9、2024-2026年3月财务报表/2025年12月利润表.xlsx',
    '/Users/jasmin0828/Downloads/蓉惠芯创基础资料/9、2024-2026年3月财务报表/2026年3月资产负债表.xlsx',
    '/Users/jasmin0828/Downloads/蓉惠芯创基础资料/9、2024-2026年3月财务报表/2026年3月利润表.xlsx',
]

print("=" * 60)
print("📊 蓉惠芯创 财务分析（带年初余额自动补充）")
print("=" * 60)

try:
    data_by_year, ind, years, trends, errors = analyze_files(files)
    print(f"\n✅ 识别期间: {years}")
    print(f"   (含从年初余额自动补充的时点)")

    for y in years:
        bs = data_by_year.get(y, {}).get('BS', {})
        is_ = data_by_year.get(y, {}).get('IS', {})
        print(f"\n  {y}: BS={len(bs)} 项, IS={len(is_)} 项")
        for k in ['货币资金', '应收账款', '资产总计', '负债合计',
                  '所有者权益合计', '营业收入', '净利润']:
            v = bs.get(k)
            if v is None:
                v = is_.get(k)
            if v is None:
                continue
            if isinstance(v, list):
                v = v[0] if v else None
            if v is not None:
                try:
                    print(f"      {k}: {v:,.2f}")
                except (TypeError, ValueError):
                    print(f"      {k}: {v}")

    print("\n" + "=" * 60)
    print("📄 生成 Excel 报告（多年对比）")
    print("=" * 60)
    try:
        generate_report(data_by_year, ind, years, trends,
                        "/tmp/蓉惠芯创_分析报告.xlsx", "蓉惠芯创")
        import os
        size = os.path.getsize("/tmp/蓉惠芯创_分析报告.xlsx")
        print(f"✅ 报告已生成: /tmp/蓉惠芯创_分析报告.xlsx ({size:,} bytes)")
        print(f"   含 {len(years)} 个时点的对比")
    except Exception as e:
        print(f"❌ 报告失败: {e}")
        import traceback
        traceback.print_exc()

    if errors:
        print("\n错误:")
        for fp, e in errors:
            print(f"  {Path(fp).name}: {e[:100]}")
except Exception as e:
    import traceback
    traceback.print_exc()
