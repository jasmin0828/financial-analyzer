"""诊断：你的 Excel 到底提取了哪些科目，还有哪些没提取到"""
from analyzer import extract_from_excel, BS_ITEMS, IS_ITEMS, CF_ITEMS
from pathlib import Path

files = [
    ('2025年资产负债表.xlsx', '/Users/jasmin0828/Downloads/蓉惠芯创基础资料/9、2024-2026年3月财务报表/2025年资产负债表.xlsx'),
    ('2025年12月利润表.xlsx', '/Users/jasmin0828/Downloads/蓉惠芯创基础资料/9、2024-2026年3月财务报表/2025年12月利润表.xlsx'),
    ('2026年3月资产负债表.xlsx', '/Users/jasmin0828/Downloads/蓉惠芯创基础资料/9、2024-2026年3月财务报表/2026年3月资产负债表.xlsx'),
    ('2026年3月利润表.xlsx', '/Users/jasmin0828/Downloads/蓉惠芯创基础资料/9、2024-2026年3月财务报表/2026年3月利润表.xlsx'),
]

print("=" * 70)
print("📊 科目提取诊断报告")
print("=" * 70)

all_found = {"BS": set(), "IS": set(), "CF": set()}

for name, path in files:
    print(f"\n━━━ {name} ━━━")
    try:
        data = extract_from_excel(path)
        for section in ['BS', 'IS', 'CF']:
            items = data.get(section, {})
            if items:
                print(f"  {section}: {len(items)} 项")
                for k, v in sorted(items.items()):
                    v_str = f"{v:,.2f}" if isinstance(v, (int, float)) else str(v)
                    print(f"    ✓ {k}: {v_str}")
                    all_found[section].add(k)
            else:
                print(f"  {section}: (无数据)")
    except Exception as e:
        print(f"  ❌ 错误: {e}")

# 对比标准科目表，看哪些没匹配
print("\n" + "=" * 70)
print("📋 标准科目 vs 实际提取（看缺什么）")
print("=" * 70)

for section, std_dict, found in [("BS", BS_ITEMS, all_found["BS"]),
                                    ("IS", IS_ITEMS, all_found["IS"]),
                                    ("CF", CF_ITEMS, all_found["CF"])]:
    print(f"\n━━━ {section} 覆盖率: {len(found)}/{len(std_dict)} 项 ({100*len(found)/len(std_dict):.0f}%) ━━━")

    missing = set(std_dict.keys()) - found
    if missing:
        print(f"  ❌ 未匹配的科目 ({len(missing)}):")
        for m in sorted(missing):
            print(f"    - {m}")
    else:
        print(f"  ✅ 全部标准科目都匹配到了！")

print("\n" + "=" * 70)
print("📈 总结")
print("=" * 70)
print(f"  BS 标准: {len(BS_ITEMS)} 项, 提取: {len(all_found['BS'])} 项")
print(f"  IS 标准: {len(IS_ITEMS)} 项, 提取: {len(all_found['IS'])} 项")
print(f"  CF 标准: {len(CF_ITEMS)} 项, 提取: {len(all_found['CF'])} 项")
print(f"  现金流文件: ❌ 没有提供（指标计算会跳过）")
