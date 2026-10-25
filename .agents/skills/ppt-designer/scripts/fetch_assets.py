#!/usr/bin/env python3
"""fetch_assets.py — 联网搜索/下载 PPT 动效素材（视频背景、动效元素），并记录版权信息。

用法:
  python3 fetch_assets.py search <关键词>            # 搜索，列出候选（标题/地址/授权）
  python3 fetch_assets.py fetch <序号|URL> --dir <assets 目录> [--name motion.mp4]

来源与授权:
  - mixkit     Mixkit Free License（免费商用、无需署名）——搜索其分类页提取 mp4 直链
  - url        直接下载给定地址（版权自负，agent 须确认来源许可后使用）

规则:
  - 下载写入 <dir>/ 并生成/合并 <dir>/license.json（source/license/fetched_at），交付时随 deck 存档
  - 单文件 >25MB 拒收；网络失败自动重试 3 次
  - WPS/PowerPoint 放映均支持内嵌 mp4；webm 需转码（本脚本不转码，仅标注）
"""
import json
import os
import re
import sys
import time
import urllib.request

UA = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) Office-Agent/1.0"}
MAX_BYTES = 25 * 1024 * 1024


def http_get(url, binary=False, retries=3, timeout=30):
    last = None
    for i in range(retries):
        try:
            req = urllib.request.Request(url, headers=UA)
            with urllib.request.urlopen(req, timeout=timeout) as r:
                data = r.read()
                return data if binary else data.decode("utf-8", "ignore")
        except Exception as e:
            last = e
            time.sleep(2 * (i + 1))
    raise RuntimeError(f"下载失败（重试 {retries} 次）: {url} — {last}")


def search_mixkit(query):
    slug = re.sub(r"[^a-z0-9]+", "-", query.strip().lower()).strip("-")
    url = f"https://mixkit.co/free-stock-video/{slug}/"
    html = http_get(url)
    seen, out = set(), []
    for m in re.finditer(r'https://assets\.mixkit\.co/videos(?:/preview)?/([^"\'\s]+?\.mp4)', html):
        u = m.group(0)
        u = re.sub(r"/preview/", "/", u)
        if u in seen:
            continue
        seen.add(u)
        title = re.sub(r"[-_]?(\d+)?(-large|-sd|-hd)?\.mp4$", "", u.rsplit("/", 1)[-1])
        title = title.replace("mixkit-", "").replace("-", " ")
        out.append({"title": title, "url": u, "source": "mixkit",
                    "license": "Mixkit Free License（免费商用，无需署名）"})
        if len(out) >= 10:
            break
    return out


def cmd_search(query):
    results = search_mixkit(query)
    if not results:
        print("无结果——换个英文关键词再试（如 abstract, particles, gradient, waves）")
        return 1
    for i, r in enumerate(results):
        print(f"[{i}] {r['title']}\n    {r['url']}\n    {r['license']}")
    return 0


def cmd_fetch(target, out_dir, name=None):
    if target.startswith("http"):
        url, meta = target, {"source": "url", "license": "未登记——使用前须确认来源许可"}
    else:
        query = os.environ.get("FA_QUERY", "")
        results = search_mixkit(query) if query else []
        try:
            idx = int(target)
            meta = results[idx]
        except (ValueError, IndexError):
            print(f"无效序号 {target}（先运行 search，或用 FA_QUERY 环境变量传关键词）")
            return 2
        url = meta["url"]
    os.makedirs(out_dir, exist_ok=True)
    ext = ".mp4" if ".mp4" in url else (".gif" if ".gif" in url else ".webm")
    path = os.path.join(out_dir, name or f"motion{ext}")
    data = http_get(url, binary=True, timeout=120)
    if len(data) > MAX_BYTES:
        print(f"拒收：文件 {len(data) / 1e6:.1f}MB 超过 25MB 上限")
        return 3
    with open(path, "wb") as f:
        f.write(data)
    lic_path = os.path.join(out_dir, "license.json")
    lic = []
    if os.path.exists(lic_path):
        try:
            lic = json.load(open(lic_path))
        except Exception:
            lic = []
    lic.append({"file": os.path.basename(path), "title": meta.get("title", ""),
                "url": url, "source": meta.get("source"),
                "license": meta.get("license"), "bytes": len(data),
                "fetched_at": time.strftime("%Y-%m-%d")})
    json.dump(lic, open(lic_path, "w"), ensure_ascii=False, indent=2)
    print(f"saved: {path} ({len(data) / 1e6:.1f}MB) — 许可已记录到 license.json")
    return 0


if __name__ == "__main__":
    if len(sys.argv) >= 3 and sys.argv[1] == "search":
        sys.exit(cmd_search(" ".join(sys.argv[2:])))
    if len(sys.argv) >= 4 and sys.argv[1] == "fetch":
        d = sys.argv[sys.argv.index("--dir") + 1] if "--dir" in sys.argv else "assets"
        nm = sys.argv[sys.argv.index("--name") + 1] if "--name" in sys.argv else None
        sys.exit(cmd_fetch(sys.argv[2], d, nm))
    print(__doc__)
    sys.exit(2)
