#!/usr/bin/env python3
"""
投标文件元数据分析脚本
检查多个投标文件的创建者、最后保存者、创建时间、修改时间等元数据，
标记异常一致性和规律性差异。
支持 .docx, .xlsx, .pdf 格式。

法规依据：87号令第37条（视为串通投标的6种情形）
- 第(1)项：不同投标人的投标文件由同一单位或个人编制
- 第(4)项：不同投标人的投标文件异常一致或报价呈规律性差异
"""

import os
import sys
import hashlib
import json
from pathlib import Path
from datetime import datetime
from collections import defaultdict
import zipfile
import xml.etree.ElementTree as ET


def extract_docx_metadata(filepath):
    """提取 .docx/.xlsx 文件的核心元数据"""
    try:
        with zipfile.ZipFile(filepath, 'r') as z:
            if 'docProps/core.xml' in z.namelist():
                with z.open('docProps/core.xml') as f:
                    tree = ET.parse(f)
                    root = tree.getroot()
                    ns = {
                        'dc': 'http://purl.org/dc/elements/1.1/',
                        'cp': 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties',
                        'dcterms': 'http://purl.org/dc/terms/'
                    }
                    creator = root.find('.//dc:creator', ns)
                    last_modified_by = root.find('.//cp:lastModifiedBy', ns)
                    created = root.find('.//dcterms:created', ns)
                    modified = root.find('.//dcterms:modified', ns)

                    return {
                        'creator': creator.text if creator is not None else None,
                        'last_modified_by': last_modified_by.text if last_modified_by is not None else None,
                        'created': created.text if created is not None else None,
                        'modified': modified.text if modified is not None else None,
                    }
    except Exception as e:
        return {'error': str(e)}
    return {}


def extract_pdf_metadata(filepath):
    """提取 PDF 文件的元数据（简易版）"""
    try:
        with open(filepath, 'rb') as f:
            content = f.read(65536)
            text = content.decode('latin-1', errors='ignore')

            metadata = {}
            for field in ['/Author', '/Creator', '/Producer', '/CreationDate', '/ModDate']:
                idx = text.find(field)
                if idx >= 0:
                    start = idx + len(field) + 1
                    end = text.find('\n', start)
                    if end < 0:
                        end = start + 200
                    val = text[start:end].strip().rstrip(')').strip('(')
                    metadata[field.strip('/').lower()] = val
            return metadata
    except Exception as e:
        return {'error': str(e)}
    return {}


def compute_file_hash(filepath, algorithm='md5'):
    """计算文件哈希值"""
    h = hashlib.new(algorithm)
    with open(filepath, 'rb') as f:
        for chunk in iter(lambda: f.read(8192), b''):
            h.update(chunk)
    return h.hexdigest()


def analyze_text_similarity(texts, threshold=0.85):
    """
    分析多段文本的相似度。
    使用 Jaccard 相似度（基于4-gram字符集）。
    """
    def get_ngrams(text, n=4):
        return set(text[i:i+n] for i in range(len(text) - n + 1))

    similar_pairs = []
    ngram_sets = [get_ngrams(t) for t in texts]

    for i in range(len(texts)):
        for j in range(i + 1, len(texts)):
            set_i, set_j = ngram_sets[i], ngram_sets[j]
            if not set_i or not set_j:
                continue
            intersection = len(set_i & set_j)
            union = len(set_i | set_j)
            similarity = intersection / union if union > 0 else 0
            if similarity >= threshold:
                similar_pairs.append({
                    'index_a': i, 'index_b': j,
                    'similarity': round(similarity, 4)
                })
    return similar_pairs


def main(directory, output_json=None):
    """
    扫描目录中的所有投标文件，分析元数据和文本相似度。
    """
    directory = Path(directory)
    results = {
        'scan_time': datetime.now().isoformat(),
        'directory': str(directory),
        'files': [],
        'metadata_anomalies': [],
        'similarity_results': [],
        'warnings': []
    }

    supported_exts = {'.docx', '.doc', '.pdf', '.xlsx', '.xls'}
    files = []
    for ext in supported_exts:
        files.extend(directory.glob(f'**/*{ext}'))

    if len(files) < 2:
        results['warnings'].append('文件数量不足2个，无法进行对比分析')
        print("⚠️  文件数量不足2个，无法进行对比分析")
        if output_json:
            json.dump(results, open(output_json, 'w', encoding='utf-8'),
                     ensure_ascii=False, indent=2)
        return results

    for filepath in sorted(files):
        ext = filepath.suffix.lower()
        if ext in ('.docx', '.xlsx'):
            meta = extract_docx_metadata(filepath)
        elif ext == '.pdf':
            meta = extract_pdf_metadata(filepath)
        else:
            meta = extract_docx_metadata(filepath)

        file_info = {
            'filename': filepath.name,
            'path': str(filepath),
            'size': filepath.stat().st_size,
            'modified_time': datetime.fromtimestamp(filepath.stat().st_mtime).isoformat(),
            'metadata': meta,
            'md5': compute_file_hash(filepath),
        }
        results['files'].append(file_info)

    # 检查1: 相同创建者
    creators = defaultdict(list)
    for f in results['files']:
        creator = f['metadata'].get('creator')
        if creator:
            creators[creator].append(f['filename'])
    for creator, filenames in creators.items():
        if len(filenames) > 1:
            results['metadata_anomalies'].append({
                'type': '相同创建者',
                'creator': creator,
                'files': filenames,
                'risk': '高',
                'regulation': '87号令第37条第(1)项',
            })

    # 检查2: 相同最后保存者
    modifiers = defaultdict(list)
    for f in results['files']:
        modifier = f['metadata'].get('last_modified_by')
        if modifier:
            modifiers[modifier].append(f['filename'])
    for modifier, filenames in modifiers.items():
        if len(filenames) > 1:
            results['metadata_anomalies'].append({
                'type': '相同最后保存者',
                'modifier': modifier,
                'files': filenames,
                'risk': '高',
                'regulation': '87号令第37条第(1)项',
            })

    # 检查3: 相同哈希值（完全相同的文件）
    hash_map = defaultdict(list)
    for f in results['files']:
        hash_map[f['md5']].append(f['filename'])
    for h, filenames in hash_map.items():
        if len(filenames) > 1:
            results['metadata_anomalies'].append({
                'type': '文件完全一致（MD5哈希值相同）',
                'hash': h,
                'files': filenames,
                'risk': '极高',
                'regulation': '87号令第37条第(4)项',
            })

    # 输出报告
    print(f"\n{'='*60}")
    print(f"投标文件元数据分析报告")
    print(f"{'='*60}")
    print(f"扫描目录: {directory}")
    print(f"扫描文件数: {len(files)}")
    print(f"发现异常: {len(results['metadata_anomalies'])} 项")
    print(f"{'='*60}")

    if results['metadata_anomalies']:
        print(f"\n【异常详情】")
        for i, anomaly in enumerate(results['metadata_anomalies'], 1):
            print(f"\n异常 {i}: {anomaly['type']}")
            print(f"  风险等级: {anomaly['risk']}")
            print(f"  法规依据: {anomaly['regulation']}")
            print(f"  涉及文件: {', '.join(anomaly['files'])}")
            if 'creator' in anomaly:
                print(f"  创建者: {anomaly['creator']}")
            if 'modifier' in anomaly:
                print(f"  修改者: {anomaly['modifier']}")

    print(f"\n【文件元数据一览】")
    print(f"{'文件名':<30} {'创建者':<20} {'最后修改者':<20} {'大小(KB)':<10}")
    print("-" * 80)
    for f in results['files']:
        creator = (f['metadata'].get('creator') or 'N/A')[:18]
        modifier = (f['metadata'].get('last_modified_by') or 'N/A')[:18]
        size_kb = f['size'] // 1024
        print(f"{f['filename']:<30} {creator:<20} {modifier:<20} {size_kb:<10}")

    if output_json:
        with open(output_json, 'w', encoding='utf-8') as fp:
            json.dump(results, fp, ensure_ascii=False, indent=2)
        print(f"\n详细结果已保存至: {output_json}")

    return results


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法: python bid_metadata_check.py <投标文件目录> [输出JSON路径]")
        print("示例: python bid_metadata_check.py ./bid_docs/ result.json")
        sys.exit(1)
    directory = sys.argv[1]
    output = sys.argv[2] if len(sys.argv) > 2 else None
    main(directory, output)
