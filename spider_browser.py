"""打开可见浏览器，手动验证后保存药品详情页面。"""

import argparse
import json
from datetime import datetime
from pathlib import Path

from playwright.sync_api import Error, TimeoutError, sync_playwright


DEFAULT_URL = (
    "https://datasearch.nmpa.gov.cn/datasearch/search-info.html?nmpa="
    "aWQ9NDA5YTM5ZjQwMDUxMWE5NzVjYzc3ZmNmNzEzYzhmNDMmaXRlbUlkPWZmODA4MDgxODNjYWQ3NTAwMTg0MDg4MWY4NDgxNzlm"
)
ROOT = Path(__file__).resolve().parent
CONTROL_HTML = """<!doctype html><html lang="zh-CN"><meta charset="utf-8">
<title>药品详情保存助手</title>
<style>body{font:18px sans-serif;max-width:760px;margin:60px auto;line-height:1.8}
button{padding:12px 24px;font-size:18px;cursor:pointer}pre{white-space:pre-wrap}</style>
<h1>药品详情保存助手</h1>
<p>请在另一个标签页手动完成滑动验证，等药品详情显示后，回到这里点击保存。
无需打开开发者工具。保存后可以再次点击，生成新的副本。</p>
<button id="save" onclick="window.captureRequested=true;this.disabled=true;
document.querySelector('#status').textContent='正在保存，请稍候……'">保存当前详情页面</button>
<pre id="status">等待手动验证。</pre>
<p>完成后关闭整个浏览器窗口即可退出脚本。</p></html>"""


def save_page(page, output_root):
    """保留渲染后的正文、网页、表格及截图，供后续解析。"""
    if page.is_closed():
        raise ValueError("详情标签页已关闭，请重新运行脚本。")
    text = page.locator("body").inner_text(timeout=10000)
    if not text.strip():
        raise ValueError("页面正文为空，请等待页面加载完成后重试。")
    tables = page.locator("table").evaluate_all("""tables => tables.map(table =>
        Array.from(table.rows, row => Array.from(row.cells, cell => cell.innerText)))""")
    output = output_root / datetime.now().strftime("%Y%m%d_%H%M%S_%f")
    output.mkdir(parents=True)
    (output / "页面正文.txt").write_text(text, encoding="utf-8")
    (output / "页面.html").write_text(page.content(), encoding="utf-8")
    (output / "表格.json").write_text(
        json.dumps({"地址": page.url, "标题": page.title(), "表格": tables},
                   ensure_ascii=False, indent=2), encoding="utf-8"
    )
    screenshot_note = ""
    try:
        page.screenshot(path=str(output / "页面.png"), full_page=True, timeout=15000)
    except Error:
        screenshot_note = "\n截图失败，正文、网页和表格已保存。"
    return (f"已保存到：{output}\n正文 {len(text)} 字，表格 {len(tables)} 个。"
            "\n请检查文件是否包含实际药品详情；保存完成不代表验证通过。" + screenshot_note)


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--url", default=DEFAULT_URL, help="要打开的药品详情地址")
    parser.add_argument("--browser", choices=["msedge", "chrome"], default="chrome",
                        help="使用已安装的浏览器，默认使用 Chrome")
    args = parser.parse_args()
    with sync_playwright() as playwright:
        browser = playwright.chromium.launch(channel=args.browser, headless=False)
        context = browser.new_context(no_viewport=True, locale="zh-CN")
        context.set_default_timeout(10000)
        control = context.new_page()
        control.set_content(CONTROL_HTML)
        page = context.new_page()
        try:
            page.goto(args.url, wait_until="domcontentloaded", timeout=45000)
        except (TimeoutError, Error):
            print("页面导航未正常完成，可在详情标签页手动刷新或重新输入地址。", flush=True)
        print("浏览器已打开，请手动验证，再到“药品详情保存助手”标签页点击保存。", flush=True)
        while browser.is_connected() and not control.is_closed():
            try:
                if control.evaluate("Boolean(window.captureRequested)"):
                    control.evaluate("window.captureRequested = false")
                    try:
                        message = save_page(page, ROOT / "browser_output")
                    except ValueError as exc:
                        message = str(exc)
                    except (Error, OSError) as exc:
                        message = f"保存未完成（异常类型：{type(exc).__name__}），请确认页面已加载并重试。"
                    print(message, flush=True)
                    control.evaluate("""message => {
                        document.querySelector('#status').textContent = message;
                        document.querySelector('#save').disabled = false;
                    }""", message)
                control.wait_for_timeout(500)
            except Error:
                break
        browser.close()
        print("浏览器会话已结束。", flush=True)


if __name__ == "__main__":
    try:
        main()
    except Error:
        print("浏览器启动或运行失败，请确认已安装 Chrome，或通过 --browser 指定已安装的浏览器。")
        raise SystemExit(1)
