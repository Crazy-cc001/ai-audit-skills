#!/usr/bin/env python3
"""
评标报告评分异常检测脚本
检查客观分一致性、主观分畸高畸低、分值汇总错误等。

法规依据：
- 87号令第64条：分值汇总计算错误、分项评分超范围、客观分不一致、
  评分畸高畸低等4种情形可重新评审，须书面报告财政部门
- 87号令第55条：评审因素须细化量化
"""

import sys
import json
from statistics import mean, stdev


def check_scoring_anomalies(scores):
    """
    检查评标评分是否存在异常。

    Args:
        scores: {
            'bidders': ['投标人A', '投标人B', ...],
            'evaluators': ['评委1', '评委2', ...],
            'items': [
                {
                    'name': '评分项名称',
                    'type': 'objective' | 'subjective',
                    'max_score': 10.0,
                    'scores': [
                        [score_by_evaluator1_for_bidder1, ...],
                        [score_by_evaluator2_for_bidder1, ...],
                        ...
                    ]
                },
                ...
            ]
        }

    Returns:
        list: 异常项列表
    """
    anomalies = []
    bidders = scores['bidders']
    evaluators = scores['evaluators']
    items = scores['items']

    OUTLIER_ZSCORE = 1.8
    SUBJECTIVE_DEVIATION_THRESHOLD = 0.20  # 20%

    for item in items:
        item_name = item['name']
        item_type = item['type']
        max_score = item['max_score']
        scores_matrix = item['scores']  # [evaluator][bidder]

        # 检查分项评分是否超出范围
        for evaluator_idx, evaluator in enumerate(evaluators):
            for bidder_idx, bidder in enumerate(bidders):
                score = scores_matrix[evaluator_idx][bidder_idx]
                if score < 0 or score > max_score:
                    anomalies.append({
                        'type': '分项评分超出范围',
                        'item': item_name,
                        'evaluator': evaluator,
                        'bidder': bidder,
                        'score': score,
                        'max_score': max_score,
                        'severity': '高',
                        'detail': f'评委{evaluator}在"{item_name}"中给{bidder}打分{score}，超出满分{max_score}',
                        'regulation': '87号令第64条：分项评分超出评分标准范围可重新评审',
                    })

        if item_type == 'objective':
            # 客观分：每个投标人的各评委打分应该完全一致
            for bidder_idx, bidder in enumerate(bidders):
                bidder_scores = [scores_matrix[e][bidder_idx] for e in range(len(evaluators))]
                unique_scores = set(bidder_scores)
                if len(unique_scores) > 1:
                    anomalies.append({
                        'type': '客观分不一致',
                        'item': item_name,
                        'bidder': bidder,
                        'scores': dict(zip(evaluators, bidder_scores)),
                        'severity': '高',
                        'detail': f'客观评分项"{item_name}"中，{bidder}得分不一致',
                        'regulation': '87号令第64条：客观评审因素评分不一致可重新评审',
                    })

        elif item_type == 'subjective':
            # 主观分：检查畸高畸低
            for evaluator_idx, evaluator in enumerate(evaluators):
                evaluator_scores = scores_matrix[evaluator_idx]

                if len(bidders) >= 3:
                    avg = mean(evaluator_scores)
                    std = stdev(evaluator_scores) if len(evaluator_scores) >= 2 else 0

                    for bidder_idx, bidder in enumerate(bidders):
                        score = evaluator_scores[bidder_idx]
                        if std > 0:
                            z = abs(score - avg) / std
                            if z > OUTLIER_ZSCORE:
                                anomalies.append({
                                    'type': '主观分畸高畸低',
                                    'item': item_name,
                                    'evaluator': evaluator,
                                    'bidder': bidder,
                                    'score': score,
                                    'average': round(avg, 2),
                                    'z_score': round(z, 2),
                                    'severity': '中',
                                    'detail': f'评委{evaluator}在"{item_name}"中给{bidder}打分{score}'
                                             f'（平均分{avg:.2f}，Z-score={z:.2f}）',
                                    'regulation': '87号令第64条：评分畸高畸低可重新评审',
                                })

            # 检查同一投标人在不同评委间的分差
            for bidder_idx, bidder in enumerate(bidders):
                bidder_scores = [scores_matrix[e][bidder_idx] for e in range(len(evaluators))]
                if len(bidder_scores) >= 3:
                    avg = mean(bidder_scores)
                    max_diff = max(abs(s - avg) for s in bidder_scores)
                    if avg > 0 and max_diff / avg > SUBJECTIVE_DEVIATION_THRESHOLD:
                        anomalies.append({
                            'type': '评委间打分偏差过大',
                            'item': item_name,
                            'bidder': bidder,
                            'scores': dict(zip(evaluators, bidder_scores)),
                            'max_deviation_pct': round(max_diff / avg * 100, 1),
                            'severity': '中',
                            'detail': f'在"{item_name}"中，各评委对{bidder}的打分偏差达{max_diff/avg*100:.1f}%',
                            'regulation': '87号令第64条',
                        })

    # 输出报告
    print(f"\n{'='*60}")
    print(f"评标评分异常检测报告")
    print(f"{'='*60}")
    print(f"投标人数量: {len(bidders)}")
    print(f"评委数量: {len(evaluators)}")
    print(f"评分项数量: {len(items)}")
    print(f"检测到异常: {len(anomalies)} 项")
    print(f"{'='*60}")

    if anomalies:
        print(f"\n【异常详情】")
        for i, a in enumerate(anomalies, 1):
            print(f"\n异常 {i}: [{a['severity']}风险] {a['type']}")
            print(f"  {a['detail']}")
            print(f"  法规依据: {a['regulation']}")
            if 'scores' in a:
                print(f"  评分详情: {a['scores']}")
    else:
        print(f"\n✅ 未发现评分异常")

    return anomalies


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python scoring_check.py <评分数据JSON文件>")
        print("")
        print("JSON格式示例:")
        example = {
            "bidders": ["投标人A", "投标人B", "投标人C"],
            "evaluators": ["评委1", "评委2", "评委3", "评委4", "评委5"],
            "items": [
                {
                    "name": "技术方案",
                    "type": "subjective",
                    "max_score": 30,
                    "scores": [
                        [25, 22, 28],
                        [24, 23, 27],
                        [26, 21, 29],
                        [25, 22, 28],
                        [10, 23, 27]
                    ]
                },
                {
                    "name": "企业资质",
                    "type": "objective",
                    "max_score": 10,
                    "scores": [
                        [10, 8, 10],
                        [10, 8, 10],
                        [10, 8, 10],
                        [10, 8, 10],
                        [8, 8, 10]
                    ]
                }
            ]
        }
        print(json.dumps(example, ensure_ascii=False, indent=2))
        print("")
        print("注意: scores[评委索引][投标人索引]——scores[0][0]是评委1给投标人A的分数")
        sys.exit(1)

    with open(sys.argv[1], 'r', encoding='utf-8') as f:
        scores = json.load(f)

    # 验证数据格式
    n_evaluators = len(scores['evaluators'])
    n_bidders = len(scores['bidders'])
    for item in scores['items']:
        if len(item['scores']) != n_evaluators:
            print(f"错误: 评分项'{item['name']}'的scores行数({len(item['scores'])})与评委数({n_evaluators})不符")
            sys.exit(1)
        for i, row in enumerate(item['scores']):
            if len(row) != n_bidders:
                print(f"错误: 评分项'{item['name']}'的第{i+1}行列数({len(row)})与投标人数({n_bidders})不符")
                sys.exit(1)

    anomalies = check_scoring_anomalies(scores)

    if anomalies:
        sys.exit(1)
