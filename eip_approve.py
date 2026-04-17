"""
EIP/BPM 簽核自動化工具
連線至已開啟的 Edge 瀏覽器（CDP port 9222）執行簽核作業。
"""

import asyncio
import argparse
import re
from playwright.async_api import async_playwright, Browser, Page

EIP_URL = "http://eip.youngoptics.com/EIP/wpa.nsf/WPAMASPG10TW?OpenPage"
CDP_ENDPOINT = "http://localhost:9222"


async def connect_browser():
    """回傳 (playwright, browser) tuple，呼叫端負責 await p.stop()。"""
    p = await async_playwright().start()
    browser = await p.chromium.connect_over_cdp(CDP_ENDPOINT)
    return p, browser


async def get_or_open_eip_page(browser: Browser) -> Page:
    """切換至已開啟的 EIP 頁面，或開啟新頁面。"""
    for ctx in browser.contexts:
        for page in ctx.pages:
            if "WPAMASPG10TW" in page.url or "eip.youngoptics.com" in page.url:
                await page.bring_to_front()
                await page.reload()
                await page.wait_for_load_state("networkidle")
                return page

    ctx = browser.contexts[0] if browser.contexts else await browser.new_context()
    page = await ctx.new_page()
    await page.goto(EIP_URL, wait_until="networkidle")
    return page


async def query_pending(browser: Browser) -> None:
    """查詢所有待簽事項並列印。"""
    page = await get_or_open_eip_page(browser)
    print("=== EIP 待簽事項 ===")

    # EIP iframe 區
    frames = page.frames
    for frame in frames:
        content = await frame.content()
        if "SelectAll" in content or "selectAllInView" in content or "待簽" in content:
            snap = await frame.inner_text("body")
            # 找出各文件類型與份數
            for line in snap.splitlines():
                line = line.strip()
                if line and ("份" in line or "件" in line or "申請" in line):
                    print(f"  {line}")

    # BPM 區（主頁面 table rows）
    print("\n=== BPM 待簽事項 ===")
    bpm_rows = await page.query_selector_all("table tr")
    found_bpm = False
    for row in bpm_rows:
        text = await row.inner_text()
        text = text.strip()
        if any(kw in text for kw in ["PUR01", "RD004", "請購", "BOM", "BPM"]):
            print(f"  {text}")
            found_bpm = True

    if not found_bpm:
        print("  （無 BPM 待簽事項）")


# ---------------------------------------------------------------------------
# EIP iframe 簽核
# ---------------------------------------------------------------------------

async def approve_eip_with_select_all(browser: Browser, doc_type: str) -> int:
    """批量核准含 Select All 的 EIP 文件（例如請假申請單）。"""
    page = await get_or_open_eip_page(browser)
    approved = 0

    for frame in page.frames:
        content = await frame.content()
        if doc_type not in content:
            continue
        select_all = await frame.query_selector("a[href*='selectAllInView']")
        if not select_all:
            continue
        print(f"找到 [{doc_type}]，執行 Select All → Approve …")
        await select_all.click()
        await page.wait_for_timeout(500)

        approve_btn = await frame.query_selector("a[href*=\"Activity('APPROVE')\"]")
        if approve_btn:
            await approve_btn.click()
            await page.wait_for_timeout(1000)
            approved += 1
            print(f"  ✅ [{doc_type}] 批量核准完成")
        break

    await page.reload()
    await page.wait_for_load_state("networkidle")
    return approved


async def approve_eip_individual(browser: Browser, doc_type: str) -> int:
    """逐筆核准無 Select All 的 EIP 文件（人力需求申請單等）。"""
    page = await get_or_open_eip_page(browser)
    approved = 0

    while True:
        page = await get_or_open_eip_page(browser)
        target_frame = None
        for frame in page.frames:
            content = await frame.content()
            if doc_type in content:
                target_frame = frame
                break
        if not target_frame:
            break

        links = await target_frame.query_selector_all("a")
        doc_link = None
        for link in links:
            href = await link.get_attribute("href") or ""
            text = await link.inner_text()
            # 文件編號連結通常是英文字母加數字（如 I2604005）
            if re.match(r"^[A-Z]\d+", text.strip()) or "OpenDocument" in href:
                doc_link = link
                break
        if not doc_link:
            break

        doc_id = (await doc_link.inner_text()).strip()
        print(f"  開啟文件 {doc_id} …")

        async with page.context.expect_page() as new_page_info:
            await doc_link.click()
        new_tab = await new_page_info.value
        await new_tab.wait_for_load_state("networkidle")

        # 選 APPROVE radio button
        approve_radio = await new_tab.query_selector(
            "input[type='radio'][value='APPROVE'], input[type='radio'][value='Approve']"
        )
        if approve_radio:
            await approve_radio.click()
            await new_tab.wait_for_timeout(300)

        # 點送核 Submit
        submit_btn = await new_tab.query_selector(
            "input[value*='Submit'], input[value*='送核'], button:text('送核'), button:text('核准')"
        )
        if submit_btn:
            await submit_btn.click()
            await new_tab.wait_for_timeout(1000)

        if not new_tab.is_closed():
            await new_tab.close()

        approved += 1
        print(f"  ✅ {doc_id} 已核准")

    return approved


async def approve_part_recognition(browser: Browser) -> int:
    """逐筆核准 TYO&KYO 物料/零件承認系統。"""
    page = await get_or_open_eip_page(browser)
    approved = 0
    doc_type = "物料/零件承認系統"

    while True:
        page = await get_or_open_eip_page(browser)
        target_frame = None
        for frame in page.frames:
            content = await frame.content()
            if doc_type in content or "零件承認" in content:
                target_frame = frame
                break
        if not target_frame:
            break

        links = await target_frame.query_selector_all("a")
        doc_link = None
        for link in links:
            text = await link.inner_text()
            href = await link.get_attribute("href") or ""
            # 承認書編號格式：R + 年月 + 流水號（如 R25010204）
            if re.match(r"^R\d+", text.strip()) or "OpenDocument" in href:
                doc_link = link
                break
        if not doc_link:
            break

        doc_id = (await doc_link.inner_text()).strip()
        print(f"  開啟零件承認 {doc_id} …")

        async with page.context.expect_page() as new_page_info:
            await doc_link.click()
        new_tab = await new_page_info.value
        await new_tab.wait_for_load_state("networkidle")

        # 選「同意承認判定結果」radio
        agree_radio = await new_tab.query_selector(
            "input[type='radio'][value*='同意'], input[type='radio'][value*='agree']"
        )
        if agree_radio:
            await agree_radio.click()
            await new_tab.wait_for_timeout(300)

        submit_btn = await new_tab.query_selector(
            "input[value*='Submit'], input[value*='送核'], button:text('送核')"
        )
        if submit_btn:
            await submit_btn.click()
            await new_tab.wait_for_timeout(1000)

        if not new_tab.is_closed():
            await new_tab.close()

        approved += 1
        print(f"  ✅ {doc_id} 零件承認已核准")

    return approved


async def approve_business_trip(browser: Browser) -> int:
    """逐筆核准 TYO/ROI 出差暨費用報支申請單（YBI-XXXXXXX）。"""
    page = await get_or_open_eip_page(browser)
    approved = 0
    doc_type = "出差暨費用報支申請單"

    while True:
        page = await get_or_open_eip_page(browser)
        target_frame = None
        for frame in page.frames:
            content = await frame.content()
            if doc_type in content or "BusinessTrip" in frame.url:
                target_frame = frame
                break
        if not target_frame:
            break

        links = await target_frame.query_selector_all("a")
        doc_link = None
        for link in links:
            text = (await link.inner_text()).strip()
            href = await link.get_attribute("href") or ""
            # 文件編號格式：YBI-XXXXXXX
            if re.match(r"^YBI-\d+", text) or "BusinessTrip" in href:
                doc_link = link
                break
        if not doc_link:
            break

        doc_id = (await doc_link.inner_text()).strip()
        print(f"  開啟出差申請單 {doc_id} …")

        async with page.context.expect_page() as new_page_info:
            await doc_link.click()
        new_tab = await new_page_info.value
        await new_tab.wait_for_load_state("networkidle")

        # 選 APPROVE radio button
        approve_radio = await new_tab.query_selector("input[type='radio']")
        if approve_radio:
            await approve_radio.click()
            await new_tab.wait_for_timeout(300)

        # 點送核 Submit
        submit_btn = await new_tab.query_selector("button:text('送核 Submit'), button:text('送核')")
        if submit_btn:
            await submit_btn.click()
            await new_tab.wait_for_timeout(1000)

        if not new_tab.is_closed():
            await new_tab.close()

        approved += 1
        print(f"  ✅ {doc_id} 出差申請單已核准")

    return approved


# ---------------------------------------------------------------------------
# BPM 簽核
# ---------------------------------------------------------------------------

def _parse_amount(text: str) -> int:
    """從字串中解析數字金額（去除逗號、貨幣符號）。"""
    nums = re.findall(r"[\d,]+", text.replace(",", ""))
    if nums:
        try:
            return int(nums[-1])
        except ValueError:
            pass
    return 0


async def _get_bpm_total_amount(form_page: Page) -> int:
    """讀取 BPM MRO 請購單的所有項目小計並加總。"""
    total = 0
    # 向下捲動確保載入全部項目
    prev_height = 0
    while True:
        await form_page.evaluate("window.scrollBy(0, 800)")
        await form_page.wait_for_timeout(300)
        height = await form_page.evaluate("document.body.scrollHeight")
        if height == prev_height:
            break
        prev_height = height

    # 找小計欄位（TWD Subtotal）
    rows = await form_page.query_selector_all("td, span, div")
    capture_next = False
    for elem in rows:
        text = (await elem.inner_text()).strip()
        if "小計" in text and "Subtotal" in text:
            capture_next = True
            continue
        if capture_next and re.search(r"[\d,]+", text):
            total += _parse_amount(text)
            capture_next = False

    # 備援：直接搜尋含數字的 subtotal 欄位
    if total == 0:
        subtotal_cells = await form_page.query_selector_all(
            "td[class*='subtotal'], td[class*='amount'], span[class*='subtotal']"
        )
        for cell in subtotal_cells:
            text = await cell.inner_text()
            total += _parse_amount(text)

    return total


async def approve_bpm_items(
    browser: Browser,
    flow_name: str,
    max_amount: int | None = None,
) -> list[dict]:
    """
    逐筆核准 BPM 待簽項目。
    flow_name: 篩選條件，如 'PUR01_MRO請購申請' 或 'RD004_Pre-BOM變更'，空字串表示全部。
    max_amount: MRO 金額上限（None 表示不限制）。
    回傳處理結果列表。
    """
    page = await get_or_open_eip_page(browser)
    results = []

    while True:
        page = await get_or_open_eip_page(browser)
        await page.wait_for_load_state("networkidle")

        # 找 BPM 列中符合條件的第一筆 Link
        rows = await page.query_selector_all("table tr")
        target_link = None
        target_meta = {}

        for row in rows:
            text = await row.inner_text()
            if flow_name and flow_name not in text:
                continue
            # 取得單號、申請人等
            cells = await row.query_selector_all("td")
            cell_texts = [((await c.inner_text()).strip()) for c in cells]

            link_elem = await row.query_selector("a")
            if link_elem:
                target_link = link_elem
                target_meta = {
                    "flow": cell_texts[0] if cell_texts else "",
                    "doc_no": cell_texts[1] if len(cell_texts) > 1 else "",
                    "applicant": cell_texts[3] if len(cell_texts) > 3 else "",
                }
                break

        if not target_link:
            break  # 無更多待簽

        doc_no = target_meta.get("doc_no", "?")
        print(f"  開啟 BPM 表單 {doc_no} …")

        async with page.context.expect_page() as new_page_info:
            await target_link.click()
        form_page = await new_page_info.value
        await form_page.wait_for_load_state("networkidle")

        # MRO 金額限制檢查
        if max_amount is not None and "PUR01" in doc_no:
            total = await _get_bpm_total_amount(form_page)
            if total > max_amount:
                print(f"  ⏭ {doc_no}｜TWD {total:,}｜超過上限 {max_amount:,}，略過")
                results.append({"doc_no": doc_no, "amount": total, "status": "skipped"})
                await form_page.close()
                # 無法繼續（此筆略過但仍在清單中，跳出避免死迴圈）
                break
            print(f"     金額 TWD {total:,} ≤ {max_amount:,}，繼續核准")
        else:
            total = 0

        # 點擊「核准/Approve」按鈕
        approve_btn = await form_page.query_selector(
            "#ext-comp-1058, button:text('核准'), button:text('Approve'), "
            "a:text('核准'), input[value*='核准'], input[value*='Approve']"
        )
        if not approve_btn:
            # 備援：找任何含「核准」文字的可點擊元素
            approve_btn = await form_page.query_selector("*:text('核准/Approve')")

        if approve_btn:
            await approve_btn.click()
            await form_page.wait_for_timeout(1500)

            # 截圖確認「提交成功」
            await form_page.screenshot(path=f"bpm_{doc_no}_confirm.png")

            # 點擊確定按鈕
            ok_btn = await form_page.query_selector(
                "#ext-element-1, button:text('確定'), button:text('OK')"
            )
            if ok_btn:
                await ok_btn.click()
                await form_page.wait_for_timeout(500)

            print(f"  ✅ {doc_no}｜{'TWD ' + str(f'{total:,}') + '｜' if total else ''}已核准")
            results.append({"doc_no": doc_no, "amount": total, "status": "approved"})
        else:
            print(f"  ⚠️ {doc_no}｜找不到核准按鈕，略過")
            results.append({"doc_no": doc_no, "amount": total, "status": "error"})

        if not form_page.is_closed():
            await form_page.close()

    return results


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

async def approve_all(browser: Browser, max_amount: int | None = None) -> None:
    """核准所有待簽事項。"""
    print("\n[1] 批量核准 EIP 請假申請單 …")
    n = await approve_eip_with_select_all(browser, "請假申請單")
    if n == 0:
        print("  無請假申請單待簽")

    print("\n[2] 逐筆核准 EIP 人力需求申請單 …")
    n = await approve_eip_individual(browser, "人力需求申請單")
    if n == 0:
        print("  無人力需求申請單待簽")

    print("\n[3] 逐筆核准 EIP 零件承認系統 …")
    n = await approve_part_recognition(browser)
    if n == 0:
        print("  無零件承認待簽")

    print("\n[3b] 逐筆核准 EIP 出差暨費用報支申請單 …")
    n = await approve_business_trip(browser)
    if n == 0:
        print("  無出差申請單待簽")

    print("\n[4] 逐筆核准 BPM MRO 請購申請 …")
    results = await approve_bpm_items(browser, "PUR01_MRO請購申請", max_amount)
    if not results:
        print("  無 MRO 請購申請待簽")
    else:
        print("\nMRO 請購申請簽核結果：")
        for r in results:
            icon = {"approved": "✅", "skipped": "⏭", "error": "⚠️"}.get(r["status"], "?")
            amt = f"TWD {r['amount']:,}｜" if r["amount"] else ""
            print(f"  {icon} {r['doc_no']}｜{amt}{r['status']}")

    print("\n[5] 逐筆核准 BPM Pre-BOM 變更 …")
    results = await approve_bpm_items(browser, "RD004_Pre-BOM變更")
    if not results:
        print("  無 Pre-BOM 變更待簽")

    print("\n完成，重新整理確認清單 …")
    page = await get_or_open_eip_page(browser)
    await page.screenshot(path="eip_after_approve.png")
    print("截圖已儲存至 eip_after_approve.png")


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

async def main() -> None:
    parser = argparse.ArgumentParser(description="EIP/BPM 簽核自動化工具")
    sub = parser.add_subparsers(dest="cmd")

    sub.add_parser("query", help="查詢待簽事項")
    sub.add_parser("approve-leave", help="核准請假申請單（批量）")
    sub.add_parser("approve-hr", help="核准人力需求申請單（逐筆）")
    sub.add_parser("approve-parts", help="核准零件承認（逐筆）")

    p_mro = sub.add_parser("approve-mro", help="核准 MRO 請購申請（逐筆）")
    p_mro.add_argument("--max-amount", type=int, default=None,
                       help="金額上限（TWD），超過此金額略過，例如 100000")

    sub.add_parser("approve-prebom", help="核准 Pre-BOM 變更（逐筆）")
    sub.add_parser("approve-trip", help="核准 TYO/ROI 出差暨費用報支申請單（逐筆）")

    p_all = sub.add_parser("approve-all", help="核准所有待簽事項")
    p_all.add_argument("--max-amount", type=int, default=None,
                       help="MRO 請購申請金額上限（TWD）")

    args = parser.parse_args()

    p, browser = await connect_browser()

    try:
        if args.cmd == "query":
            await query_pending(browser)

        elif args.cmd == "approve-leave":
            n = await approve_eip_with_select_all(browser, "請假申請單")
            print(f"完成，共批量核准 {n} 批次")

        elif args.cmd == "approve-hr":
            n = await approve_eip_individual(browser, "人力需求申請單")
            print(f"完成，共核准 {n} 筆人力需求申請單")

        elif args.cmd == "approve-parts":
            n = await approve_part_recognition(browser)
            print(f"完成，共核准 {n} 筆零件承認")

        elif args.cmd == "approve-mro":
            results = await approve_bpm_items(
                browser, "PUR01_MRO請購申請", getattr(args, "max_amount", None)
            )
            approved = sum(1 for r in results if r["status"] == "approved")
            print(f"完成，共核准 {approved}/{len(results)} 筆 MRO 請購申請")

        elif args.cmd == "approve-prebom":
            results = await approve_bpm_items(browser, "RD004_Pre-BOM變更")
            print(f"完成，共核准 {len(results)} 筆 Pre-BOM 變更")

        elif args.cmd == "approve-trip":
            n = await approve_business_trip(browser)
            print(f"完成，共核准 {n} 筆出差申請單")

        elif args.cmd == "approve-all":
            await approve_all(browser, getattr(args, "max_amount", None))

        else:
            parser.print_help()
    finally:
        # 斷開 CDP 連線（不關閉瀏覽器），再停止 playwright 子程序
        await browser.close()
        await p.stop()


if __name__ == "__main__":
    asyncio.run(main())
