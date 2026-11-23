#!/usr/bin/env python3
"""
异常低价检测脚本
根据安徽省规定自动检测投标报价是否触发异常低价评审阈值：
- 房屋建筑：低于招标控制价的 90%
- 轨道交通：低于招标控制价的 88%
- 市政基础设施：低于招标控制价的 85%

法规依据：建市函〔2020〕1073号、建市函〔2022〕453号
"""

import sys
import json
from datetime import datetime


def check_abnormal_low_price(control_price, bid_prices, project_type='housing'):
    """
    检查是否存在异常低价投标。

    Args:
        control_price: 招标控制价（最高投标限价）
        bid_prices: 各投标人报价列表 [{'name': str, 'price': float}, ...]
        project_type: 'housing'(房建), 'rail'(轨道交通), 'municipal'(市政)

    Returns:
        dict: 检查结果
    """
    thresholds = {
        'housing': (0.90, '房屋建筑'),
        'rail': (0.88, '轨道交通'),
        'municipal': (0.85, '市政基础设施'),
    }

    if project_type not in thresholds:
        return {'error': f'不支持的项目类型，支持: {list(thresholds.keys())}'}

    ratio, type_name = thresholds[project_type]
    threshold_price = control_price * ratio

    results = {
        'check_time': datetime.now().isoformat(),
        'project_type': type_name,
        'control_price': control_price,
        'threshold_ratio': f'{ratio*100}%',
        'threshold_price': round(threshold_price, 2),
        'abnormal_bids': [],
        'normal_bids': [],
    }

    for bid in bid_prices:
        bid['ratio_to_control'] = round(bid['price'] / control_price * 100, 2)
        if bid['price'] < threshold_price:
            bid['deviation'] = round(threshold_price - bid['price'], 2)
            results['abnormal_bids'].append(bid)
        else:
            results['normal_bids'].append(bid)

    print(f"\n{'='*60}")
    print(f"异常低价检测报告")
    print(f"{'='*60}")
    print(f"项目类型: {type_name}")
    print(f"异常低价阈值: 低于招标控制价的 {ratio*100}%")
    print(f"招标控制价: {control_price:,.2f} 元")
    print(f"异常低价线: {threshold_price:,.2f} 元")
    print(f"{'='*60}")

    if results['abnormal_bids']:
        print(f"\n⚠️  发现 {len(results['abnormal_bids'])} 个异常低价投标:")
        for bid in sorted(results['abnormal_bids'], key=lambda x: x['price']):
            print(f"  - {bid['name']}: {bid['price']:,.2f} 元 "
                  f"({bid['ratio_to_control']}% - 低于阈值{ratio*100}%)")
        print(f"\n法规要求：")
        print(f"  1. 评标委员会应在推荐中标候选人前要求投标人书面澄清说明并提供证明材料")
        print(f"  2. 对合同履行能力及工程质量安全风险进行全面评估")
        print(f"  3. 评估结果作为评标结果附件提交招标人")
        print(f"  4. 不能有效澄清或认定存在风险的，应否决其投标")
        print(f"\n参考法规: 建市函〔2020〕1073号、建市函〔2022〕453号")
    else:
        print(f"\n✅ 所有投标报价均不低于异常低价线")

    return results


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python abnormal_price_check.py <招标控制价> <项目类型> [投标报价列表] 或 <JSON文件>")
        print("项目类型: housing(房建), rail(轨道交通), municipal(市政)")
        print("")
        print("方式1 - 命令行参数:")
        print("  python abnormal_price_check.py 10000000 housing 8900000,9200000,9500000")
        print("")
        print("方式2 - JSON文件:")
        print("  python abnormal_price_check.py bids.json")
        print("  JSON格式: {\"control_price\": 10000000, \"project_type\": \"housing\",")
        print("            \"bids\": [{\"name\": \"A公司\", \"price\": 8900000}, ...]}")
        sys.exit(1)

    # 尝试JSON文件模式
    if sys.argv[1].endswith('.json'):
        with open(sys.argv[1], 'r', encoding='utf-8') as f:
            data = json.load(f)
        control_price = data['control_price']
        project_type = data.get('project_type', 'housing')
        bid_prices = data['bids']
    else:
        control_price = float(sys.argv[1])
        project_type = sys.argv[2]
        if len(sys.argv) >= 4:
            prices = sys.argv[3].split(',')
            bid_prices = []
            for i, price_str in enumerate(prices):
                bid_prices.append({
                    'name': f'投标人{i+1}',
                    'price': float(price_str.strip())
                })
        else:
            print("请提供投标报价列表或JSON文件")
            sys.exit(1)

    results = check_abnormal_low_price(control_price, bid_prices, project_type)

    if results.get('abnormal_bids'):
        sys.exit(1)  # 发现异常低价，退出码1
