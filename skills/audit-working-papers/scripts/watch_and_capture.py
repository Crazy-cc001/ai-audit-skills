# -*- coding: utf-8 -*-
r"""
监控新纪元软件年度切换, 自动快照MyTemp0并提取+核对
用法: python watch_and_capture.py
日志: C:\path\to\temp\xsj_pt\watch_log.txt
状态: C:\path\to\temp\xsj_pt\watch_status.txt
"""
import os, re, sys, time, shutil, subprocess, json

sys.path.insert(0, r'C:\path\to\temp\xsj_pt\pypx\src\pypxlib-2.5')

XSJ_INI = r'D:\新纪元\Xsjzb\xsjcd.ini'
MYTEMP = r'D:\新纪元\Xsjzb\MyTemp0'
BASE = r'C:\path\to\temp\xsj_pt'
LOG = os.path.join(BASE, 'watch_log.txt')
STATUS = os.path.join(BASE, 'watch_status.txt')
TARGETS = [str(y) for y in range(2016, 2026)]   # 16-25

def log(msg):
    line = time.strftime('%H:%M:%S ') + msg
    print(line, flush=True)
    with open(LOG, 'a', encoding='utf-8') as f:
        f.write(line + '\n')

def read_ini_year():
    """共享读 xsjcd.ini 取 年度 与 帐套编号"""
    try:
        with open(XSJ_INI, 'rb') as f:
            data = f.read()
        txt = data.decode('gbk', 'replace')
        m1 = re.search(r'年度=(\d{4})', txt)
        m2 = re.search(r'帐套编号=(\S+)', txt)
        m3 = re.search(r'帐套名称=(\S+)', txt)
        return (m1.group(1) if m1 else None,
                m2.group(1) if m2 else None,
                m3.group(1) if m3 else None)
    except Exception:
        return (None, None, None)

def snapshot_dir(year, ztbh):
    dst = os.path.join(BASE, f'snap{year}')
    os.makedirs(dst, exist_ok=True)
    # 清空旧快照, 重新复制
    for f in os.listdir(dst):
        try:
            os.remove(os.path.join(dst, f))
        except OSError:
            pass
    n_ok = 0
    try:
        files = [f for f in os.listdir(MYTEMP) if os.path.isfile(os.path.join(MYTEMP, f))]
    except FileNotFoundError:
        return 0
    for name in files:
        src = os.path.join(MYTEMP, name)
        try:
            with open(src, 'rb') as f:
                data = f.read()
            with open(os.path.join(dst, name), 'wb') as f:
                f.write(data)
            n_ok += 1
        except OSError:
            pass
    return n_ok

def run_extract(year):
    r = subprocess.run([sys.executable, os.path.join(BASE, 'extract_year.py'),
                        os.path.join(BASE, f'snap{year}'), str(year),
                        os.path.join(BASE, f'out{year}')],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    tail = (r.stdout or '').strip().splitlines()
    jf = [l for l in tail if '借合计' in l]
    return (r.returncode, (jf[0] if jf else '?'))

def run_recon(year):
    # 核对脚本按年度参数化: 用通用脚本 recon_year.py
    r = subprocess.run([sys.executable, os.path.join(BASE, 'recon_year.py'), str(year)],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    return r.stdout or r.stderr or ''

def run_fill(year):
    # 将年度数据填入底稿对应年列
    r = subprocess.run([sys.executable, os.path.join(BASE, 'fill_year.py'), str(year)],
                       capture_output=True, text=True, encoding='utf-8', errors='replace')
    return (r.stdout or r.stderr or '').strip().replace('\n', '; ')

done = {}
if os.path.exists(STATUS):
    try:
        for line in open(STATUS, encoding='utf-8'):
            parts = line.strip().split('\t')
            if len(parts) == 3:
                done[parts[0]] = parts[2]
    except OSError:
        pass

log('监控启动, 已完成: ' + str(done))
last_key = None
while True:
    y, zt, mc = read_ini_year()
    key = (y, zt)
    if key != last_key:
        last_key = key
        log(f'检测到当前账套: 年度={y} 编号={zt} 名称={mc}')
        if y in TARGETS and y not in done:
            time.sleep(8)   # 等软件解包完成
            n = snapshot_dir(y, zt)
            log(f'{y}年 快照复制 {n} 个文件')
            if n >= 100:
                rc, bal = run_extract(y)
                log(f'{y}年 提取完成 rc={rc} {bal}')
                rpt = run_recon(y)
                for line in rpt.strip().splitlines():
                    if '差异' in line or '相符' in line or '!!' in line:
                        log(f'  [核对] {line}')
                fillmsg = run_fill(y)
                log(f'  [填充] {fillmsg}')
                done[y] = f'{time.strftime("%H:%M")} {bal}'
                with open(STATUS, 'w', encoding='utf-8') as f:
                    for k, v in sorted(done.items()):
                        f.write(f'{k}\t{zt}\t{v}\n')
            else:
                log(f'{y}年 快照文件数不足({n}), 可能切换未完成, 稍后重试')
                done.pop(y, None)
    time.sleep(4)
