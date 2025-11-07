import csv
import os
import sys
import asyncio
import argparse
import time
import multiprocessing
import psutil
from datetime import datetime, timedelta
from playwright.async_api import async_playwright

from crawler.web_crawler import WebCrawlerAgent
from reporter.report_generation_mp import ReportGenerationAgent
from utils.extract_problematic_links import extract_error_links_from_json


def load_websites(path: str):
    """載入網站設定檔"""
    websites_config = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)
        for row in reader:
            websites_config.append(row)
    return websites_config


async def _async_crawl_worker(site_config: dict) -> dict:
    """
    Subprocess 中 asyncio 迴圈內執行的真正爬蟲
    """
    url = site_config["URL"]
    name = site_config.get("name", "")
    depth = site_config["global_depth"]
    save_html = site_config["global_save_html"]
    enable_pagination = site_config["global_enable_pagination"]

    print(f"\n🔍 [PID {os.getpid()}] 開始處理網站: {name or url}")
    
    try:
        # 在子進程中建立 crawler 和 playwright  
        async with async_playwright() as p:
            browser = await p.chromium.launch()
            crawler = WebCrawlerAgent(save_html_files=save_html, enable_pagination=enable_pagination)
            
            start_time = time.time()
            
            # 1. 執行爬蟲
            crawl_results = await crawler.crawl_site(browser, url, name=name, max_depth=depth)
            crawl_duration = time.time() - start_time
            crawl_duration_formatted = f"{int(crawl_duration // 60)}分{int(crawl_duration % 60)}秒"
            
            # 2. 提取大型結果資料
            page_summary = crawler.get_page_summary()
            external_link_results = crawler.get_external_link_results()

            # 3. 儲存 JSON/Log
            json_path = crawler.save_page_summary_to_json()
            if json_path:
                extract_error_links_from_json(json_path)
            crawler.save_crawl_log()

            # 4. 預先計算統計數據給 Excel
            one_year_ago = datetime.now() - timedelta(days=365)
            total_pages = len(crawl_results)
            failed_pages = sum(1 for status in crawl_results if status >= 400 or status == 0)
            failed_external_links = sum(1 for link_info in external_link_results.values() 
                                       if link_info.get('status', 0) >= 400 or link_info.get('status', 0) == 0)
            
            # 計算日期相關統計
            today = datetime.now().date()
            pages_with_date = 0
            no_date_pages = 0
            outdated_pages = 0
            past_dates = []  # 今天或以前的日期
            future_dates = []  # 未來日期
            
            for url_key, page_info in page_summary.items():
                last_updated = page_info.get('last_updated', '')
                
                # 統計無日期的頁面
                if last_updated == "[無日期]" or last_updated == "[爬取失敗]" or not last_updated:
                    no_date_pages += 1
                    continue
                
                try:
                    update_date = datetime.strptime(last_updated, '%Y-%m-%d')
                    update_date_only = update_date.date()
                    
                    if update_date_only <= today:
                        past_dates.append(update_date)
                        # 檢查是否為一年前的內容
                        if update_date < one_year_ago:
                            outdated_pages += 1
                    else:
                        future_dates.append(update_date)
                        
                except ValueError:
                    no_date_pages += 1
                    continue
            
            # 計算最新更新日期：優先使用過去日期的最新值，沒有才用最接近今天的未來日期
            if past_dates:
                latest_update = max(past_dates).strftime('%Y-%m-%d')
            elif future_dates:
                latest_update = min(future_dates).strftime('%Y-%m-%d')
            else:
                latest_update = "無有效日期"
            
            # 計算一年前內容的比例
            pages_with_date = len(past_dates) + len(future_dates)
            outdated_percentage = (outdated_pages / pages_with_date * 100) if pages_with_date > 0 else 0
            
            # 5. 建立小型結果字典
            stats_for_excel = {
                'site_name': name or url,
                'site_url': url,
                'total_pages': total_pages,
                'pages_with_date': pages_with_date,
                'no_date_pages': no_date_pages,
                'latest_update': latest_update,
                'outdated_pages': outdated_pages,
                'outdated_percentage': round(outdated_percentage, 2),
                'failed_pages': failed_pages,
                'failed_external_links': failed_external_links,
                'total_external_links': len(external_link_results),
                'crawl_duration': crawl_duration_formatted
            }
            
            # 6. 清理並關閉
            del page_summary
            del external_link_results
            await crawler.close()
            crawler.clear_memory()
            del crawler
            await browser.close()
            
            print(f"✅ [PID {os.getpid()}] 網站 '{name or url}' 處理完成")
            return stats_for_excel
                
    except Exception as e:
        print(f"❌ [PID {os.getpid()}] 處理網站 '{name or url}' 時發生錯誤: {e}")
        
        # 發生錯誤時也要進行清理和等待
        try:
            if 'crawler' in locals() and crawler:
                await crawler.close()
                crawler.clear_memory()
                del crawler
            if 'browser' in locals() and browser:
                await browser.close()
            
            # 給予 5 秒緩衝時間
            await asyncio.sleep(5)
            
        except Exception as cleanup_e:
            print(f"💥 [PID {os.getpid()}] 在錯誤清理中發生了額外錯誤: {cleanup_e}")
            
        return None  # 發生錯誤時返回 None


def run_crawl_task(site_config: dict) -> dict:
    """
    multiprocessing.Pool 呼叫的包裝函數
    它會建立自己的 asyncio 迴圈
    並且自我監控記憶體，超標時自動回收
    """
    
    # 讀取主進程傳來的記憶體上限
    max_mem_mb = site_config.get('global_max_mem_mb', 1024)  # 預設 1024MB (1GB)
    
    # 檢查自己的記憶體用量
    try:
        process = psutil.Process(os.getpid())
        memory_mb = process.memory_info().rss / 1024 / 1024
        
        if memory_mb > max_mem_mb:
            print(f"♻️ [PID {os.getpid()}] 記憶體超標 ({memory_mb:.1f} MB > {max_mem_mb} MB)，自動回收 worker process")
            sys.exit()  # 自殺，讓 Pool 啟動新的乾淨進程
            
    except Exception as e:
        print(f"⚠️ [PID {os.getpid()}] 記憶體檢查失敗: {e}")

    # 沒超標，才執行爬蟲任務
    return asyncio.run(_async_crawl_worker(site_config))


def auto_shutdown_vm():
    """
    自動關閉 GCE VM 執行個體
    """
    try:
        import subprocess
        
        # 直接使用固定的 VM 名稱和區域
        vm_name = "crawler-webcheck-mpstandard"
        zone = "asia-east1-c"
        
        print(f"🎉 任務全部完成，準備自動關閉 VM: {vm_name}")
        print(f"📍 VM 位置: {zone}")
        
        # 執行關機指令
        shutdown_cmd = f"gcloud compute instances stop {vm_name} --zone={zone} --quiet"
        print(f"💻 執行關機指令: {shutdown_cmd}")
        
        shutdown_result = subprocess.run(
            shutdown_cmd.split(),
            capture_output=True, text=True, timeout=60
        )
        
        if shutdown_result.returncode == 0:
            print("✅ VM 關機指令執行成功")
        else:
            print(f"❌ VM 關機指令執行失敗: {shutdown_result.stderr}")
            
    except Exception as e:
        print(f"⚠️ 自動關機失敗: {e}")
        print("ℹ️ VM 將保持開啟狀態")


def main():
    """
    主函數
    """
    parser = argparse.ArgumentParser(description='網站爬蟲和分析工具 (Multiprocessing 版本)')
    parser.add_argument('--depth', type=int, default=2, 
                       help='爬蟲的最大深度 (預設: 2)')
    parser.add_argument('--config', type=str, default="config/websites.csv",
                       help='網站設定檔案路徑 (預設: config/websites.csv)')
    parser.add_argument('--concurrent', type=int, default=2,
                       help='同時處理的網站數量 (預設: 2)')
    parser.add_argument('--no-save-html', action='store_true',
                       help='不儲存HTML檔案，僅產生統計報告 (提升效能，節省磁碟空間)')
    parser.add_argument('--no-pagination', action='store_true',
                       help='禁用分頁爬取，將有分頁參數的頁面視為重複頁面跳過 (提升效能)')
    parser.add_argument('--max-mem-mb', type=int, default=1024,
                       help='subprocess 記憶體上限 (MB)，超過此值將自動回收 (預設: 1024)')
    
    args = parser.parse_args()
    
    # 取得全域預設值
    global_depth = args.depth
    global_save_html = not args.no_save_html
    global_enable_pagination = not args.no_pagination
    
    if not os.path.exists(args.config):
        print(f"錯誤：找不到設定檔案 {args.config}")
        sys.exit(1)
    
    websites = load_websites(args.config)
    print(f"載入了 {len(websites)} 個網站，最大爬蟲深度: {global_depth}，並行數量: {args.concurrent}")
    if global_save_html:
        print("💾 HTML檔案儲存: 啟用")
    else:
        print("🚀 HTML檔案儲存: 停用 (僅產生統計報告，提升效能)")
    
    if global_enable_pagination:
        print("📄 分頁爬取: 啟用")
    else:
        print("⚡ 分頁爬取: 停用 (分頁視為重複頁面跳過，提升效能)")

    print("🚀 啟動 Multiprocessing 網站爬蟲...")

    # 1. Mainprocess 初始化 Reporter
    reporter = ReportGenerationAgent()
    output_path = reporter.initialize_excel_report()
    print(f"Excel 報告檔案初始化完成: {output_path}")
    
    processed_urls = reporter.get_processed_urls()
    
    # 2. 準備要傳遞給子進程的任務列表
    websites_to_process = []
    for site in websites:
        url = site["URL"]
        if url.strip() in processed_urls:
            continue
            
        # 參數合併處理  
        try:
            # 讀取表格內的 depth 值
            csv_depth = int(site.get('depth')) if site.get('depth') else None
            # 只有當全域深度大於表格內深度時，才使用表格內的較小值
            if csv_depth is not None and global_depth > csv_depth:
                site_depth = csv_depth
            else:
                site_depth = global_depth
        except (ValueError, TypeError):
            site_depth = global_depth
        site['global_depth'] = site_depth
        
        # 嘗試讀取 'save_html'
        if site.get('save_html', '').lower() == 'true':
            site['global_save_html'] = True
        elif site.get('save_html', '').lower() == 'false':
            site['global_save_html'] = False
        else:
            site['global_save_html'] = global_save_html  # CSV 中為空，使用全域設定
        
        # 嘗試讀取 'pagination'
        if site.get('pagination', '').lower() == 'true':
            site['global_enable_pagination'] = True
        elif site.get('pagination', '').lower() == 'false':
            site['global_enable_pagination'] = False
        else:
            site['global_enable_pagination'] = global_enable_pagination  # CSV 中為空，使用全域設定
        
        site['global_max_mem_mb'] = args.max_mem_mb
            
        websites_to_process.append(site)
    
    print(f"📋 總共 {len(websites)} 個網站，剩餘 {len(websites_to_process)} 個待處理")
    
    if not websites_to_process:
        print("🎉 所有網站都已處理完成！")
        reporter.finalize_excel_report()
        print(f"📄 報告已儲存到: {output_path}")
        
        # 關機邏輯
        auto_shutdown_vm()
        return

    # 3. 使用 multiprocessing.Pool + memory 管理
    print(f"\n🚀 啟動 {args.concurrent} 個並行處理程序，每個記憶體上限 {args.max_mem_mb} MB")
    start_time = time.time()
    
    crawl_success = True
    successful_sites = 0
    failed_sites = 0
    
    try:
        # 建立一個 Pool
        with multiprocessing.Pool(processes=args.concurrent) as pool:

            # 使用 imap_unordered 來即時取得 worker 結果
            results = pool.imap_unordered(run_crawl_task, websites_to_process)
            
            # Main process 接收從 sub process 傳回的「小字典」
            for stats_for_excel in results:
                if stats_for_excel:
                    # Main process 呼叫 add_site_to_excel 寫入結果
                    try:
                        # 在 Main process 中記錄爬取日期
                        crawl_date = datetime.now().strftime('%Y-%m-%d %H:%M')
                        stats_for_excel['crawl_date'] = crawl_date
                        
                        reporter.add_site_to_excel(stats_for_excel)
                        successful_sites += 1
                    except Exception as e:
                        print(f"❌ 寫入 Excel 失敗: {e}")
                        failed_sites += 1
                else:
                    failed_sites += 1

        total_duration = time.time() - start_time
        total_duration_formatted = f"{int(total_duration // 60)}分{int(total_duration % 60)}秒"
        
        print(f"\n🎉 並行處理完成!")
        print(f"📊 成功處理: {successful_sites} 個網站")
        print(f"❌ 失敗: {failed_sites} 個網站") 
        print(f"⏱️ 總耗時: {total_duration_formatted}")

    except Exception as e:
        print(f"\n💥 處理過程中發生嚴重錯誤: {e}")
        crawl_success = False

    finally:
        # Main process 完成報告
        reporter.finalize_excel_report()
        print(f"\n📄 報告已儲存到: {output_path}")
        print(f"✅ 總共處理了 {len(websites_to_process)} 個新網站")
        
        # Main process 呼叫關機
        if crawl_success:
            print("準備關機...")
            # 直接呼叫同步函數
            auto_shutdown_vm()
        else:
            print("🔧 由於執行過程中發生錯誤，VM 將保持開啟狀態以便除錯")


if __name__ == "__main__":
    # 確保 multiprocessing 在 macOS/Windows 上正常運作
    multiprocessing.freeze_support() 
    main()  # 直接呼叫同步的 main