#!/usr/bin/env python3
"""
招标投标关键时间节点合规性检查脚本
检查各环节时间间隔是否符合法定要求。

法规依据：
- 政府采购法第35条：招标文件发出至投标截止不得少于20日
- 政府采购法第46条：中标通知书发出后30日内签订合同
- 实施条例第31条：招标文件发售期不得少于5个工作日
- 87号令第38条：保证金退还时限
- 87号令第58条：评标报告提交时限
- 87号令第69条：中标结果公告时限
- 实施条例第50条：合同公告时限
- 建市函〔2022〕453号：招标计划提前30日发布
- 安徽301号令：中标通知书7日内发出
"""

import sys
import json
from datetime import datetime


def check_timeline(events):
    """
    检查招标投标各环节时间节点合规性。

    Args:
        events: 事件列表，每项包含 {'event': str, 'date': str (YYYY-MM-DD)}

    Returns:
        list: 违规项列表
    """
    violations = []

    dates = {}
    for e in events:
        dates[e['event']] = datetime.strptime(e['date'], '%Y-%m-%d')

    checks = [
        {
            'name': '招标文件发售期',
            'start': '招标公告发布日',
            'end': '招标文件发售截止日',
            'min_days': 4,  # ≥5个工作日，按自然日近似
            'regulation': '实施条例第31条：招标文件发售期不得少于5个工作日',
            'required': True,
        },
        {
            'name': '投标截止时间',
            'start': '招标文件发出日',
            'end': '投标截止日',
            'min_days': 19,  # ≥20日
            'regulation': '政府采购法第35条：招标文件发出至投标截止不得少于20日',
            'required': True,
        },
        {
            'name': '中标通知书发出',
            'start': '中标人确定日',
            'end': '中标通知书发出日',
            'max_days': 7,
            'regulation': '安徽301号令：中标通知书应在确定中标人后7日内发出',
            'required': False,
        },
        {
            'name': '评标报告提交',
            'start': '评标结束日',
            'end': '评标报告提交日',
            'max_days': 2,
            'regulation': '87号令第58条：代理机构评审结束后2个工作日内送评审报告',
            'required': False,
        },
        {
            'name': '中标结果公告',
            'start': '中标人确定日',
            'end': '中标结果公告日',
            'max_days': 2,
            'regulation': '87号令第69条：确定中标人后2个工作日内公告',
            'required': False,
        },
        {
            'name': '合同签订',
            'start': '中标通知书发出日',
            'end': '合同签订日',
            'max_days': 30,
            'regulation': '政府采购法第46条：中标通知书发出后30日内签订合同',
            'required': True,
        },
        {
            'name': '保证金退还（未中标）',
            'start': '中标通知书发出日',
            'end': '未中标保证金退还日',
            'max_days': 5,
            'regulation': '87号令第38条：中标通知书发出后5个工作日内退还未中标人保证金',
            'required': False,
        },
        {
            'name': '合同公告',
            'start': '合同签订日',
            'end': '合同公告日',
            'max_days': 2,
            'regulation': '实施条例第50条：合同签订后2个工作日内公告',
            'required': False,
        },
        {
            'name': '招标计划发布',
            'start': '招标计划发布日',
            'end': '招标公告发布日',
            'min_days': 29,  # ≥30日
            'regulation': '建市函〔2022〕453号：招标计划应在招标公告前不少于30日发布',
            'required': False,
        },
    ]

    for check in checks:
        start_event = check.get('start')
        end_event = check.get('end')

        if start_event not in dates or end_event not in dates:
            if check['required']:
                violations.append({
                    'check': check['name'],
                    'status': '无法检查（缺少日期信息）',
                    'regulation': check['regulation'],
                    'required': True,
                })
            continue

        start_date = dates[start_event]
        end_date = dates[end_event]
        actual_days = (end_date - start_date).days

        if actual_days < 0:
            violations.append({
                'check': check['name'],
                'status': '日期顺序错误',
                'detail': f'{end_event}({end_date.date()})早于{start_event}({start_date.date()})',
                'regulation': check['regulation'],
            })
            continue

        if 'min_days' in check:
            if actual_days < check['min_days']:
                violations.append({
                    'check': check['name'],
                    'status': '不合规',
                    'required_min': f'{check["min_days"]+1}日',
                    'actual': f'{actual_days}日',
                    'gap': f'缺少 {check["min_days"] + 1 - actual_days} 日',
                    'regulation': check['regulation'],
                })
        elif 'max_days' in check:
            if actual_days > check['max_days']:
                violations.append({
                    'check': check['name'],
                    'status': '不合规',
                    'required_max': f'{check["max_days"]}日/工作日',
                    'actual': f'{actual_days}日',
                    'gap': f'超出 {actual_days - check["max_days"]} 日',
                    'regulation': check['regulation'],
                })

    # 输出报告
    print(f"\n{'='*60}")
    print(f"招标投标关键时间节点合规性检查报告")
    print(f"{'='*60}")

    for e in events:
        print(f"  {e['event']}: {e['date']}")

    print(f"\n共检查 {len(checks)} 项时间节点要求")

    if violations:
        print(f"\n⚠️  发现 {len(violations)} 项不合规:")
        for i, v in enumerate(violations, 1):
            print(f"\n  {i}. {v['check']} - {v['status']}")
            if 'required_min' in v:
                print(f"     合规要求: 不少于 {v['required_min']}")
            elif 'required_max' in v:
                print(f"     合规要求: 不超过 {v['required_max']}")
            print(f"     实际情况: {v.get('actual', 'N/A')}")
            if 'gap' in v:
                print(f"     偏差: {v['gap']}")
            print(f"     法规依据: {v['regulation']}")
    else:
        print(f"\n✅ 所有可检查的时间节点均合规")

    return violations


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python timeline_check.py <事件JSON文件>")
        print("")
        print("JSON格式示例:")
        example = [
            {"event": "招标计划发布日", "date": "2026-01-01"},
            {"event": "招标公告发布日", "date": "2026-02-01"},
            {"event": "招标文件发出日", "date": "2026-02-01"},
            {"event": "招标文件发售截止日", "date": "2026-02-06"},
            {"event": "投标截止日", "date": "2026-02-21"},
            {"event": "评标结束日", "date": "2026-02-22"},
            {"event": "评标报告提交日", "date": "2026-02-24"},
            {"event": "中标人确定日", "date": "2026-03-01"},
            {"event": "中标通知书发出日", "date": "2026-03-05"},
            {"event": "合同签订日", "date": "2026-04-01"},
            {"event": "合同公告日", "date": "2026-04-03"},
            {"event": "未中标保证金退还日", "date": "2026-03-10"},
        ]
        print(json.dumps(example, ensure_ascii=False, indent=2))
        sys.exit(1)

    with open(sys.argv[1], 'r', encoding='utf-8') as f:
        events = json.load(f)

    violations = check_timeline(events)

    if violations:
        sys.exit(1)
