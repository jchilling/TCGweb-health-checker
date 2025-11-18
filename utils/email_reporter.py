import os
import smtplib
import zipfile
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime


class EmailReporter:
    def __init__(self):
        """從 .env 載入 Email 憑證並設定"""
        self.GMAIL_USER = os.getenv("GMAIL_USER")
        self.GMAIL_APP_PASSWORD = os.getenv("GMAIL_APP_PASSWORD")
        self.TO_EMAIL = os.getenv("TO_EMAIL", self.GMAIL_USER)
        
        # 20MB 限制 (最大25MB)
        self.MAX_ZIP_SIZE_BYTES = 20 * 1024 * 1024 
    
        if not self.GMAIL_USER or not self.GMAIL_APP_PASSWORD:
            print("❌ 錯誤：EmailReporter 未能從 .env 載入 GMAIL_USER 或 GMAIL_APP_PASSWORD")
            self.valid = False
        else:
            self.valid = True
            print(f"📧 EmailReporter 已初始化 - 發送者: {self.GMAIL_USER}, 收件者: {self.TO_EMAIL}")

    def _send_part(self, zip_filename: str, part_num: int, total_parts: int, files_in_zip: list):
        """
        發送單一的 .zip 附件
        """
        if not self.valid:
            return False
        
        print(f"📧 準備發送 Email (Part {part_num}/{total_parts}) - 檔案: {zip_filename}")
        
        try:
            # 檢查檔案大小
            zip_size_mb = os.path.getsize(zip_filename) / 1024 / 1024
            
            msg = MIMEMultipart()
            msg['Subject'] = f"網站爬蟲數據包 Part {part_num}/{total_parts} - {datetime.now().strftime('%Y-%m-%d')}"
            msg['From'] = self.GMAIL_USER
            msg['To'] = self.TO_EMAIL
            
            body_files = "\n".join([f"- {f}" for f in files_in_zip[:10]])  # 只顯示前10個
            if len(files_in_zip) > 10:
                body_files += f"\n... 及其他 {len(files_in_zip) - 10} 個檔案"
                
            body = (
                f"爬蟲任務已完成。\n\n"
                f"這是 {total_parts} 封郵件中的第 {part_num} 封。\n\n"
                f"此壓縮檔包含以下內容：\n{body_files}\n\n"
                f"壓縮檔大小: {zip_size_mb:.2f} MB\n\n"
                "(這是由 GCP VM 自動發送)"
            )
            msg.attach(MIMEText(body, 'plain'))
            
            with open(zip_filename, "rb") as f:
                part = MIMEApplication(f.read(), Name=zip_filename)
            part['Content-Disposition'] = f'attachment; filename="{zip_filename}"'
            msg.attach(part)
            
            with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
                server.login(self.GMAIL_USER, self.GMAIL_APP_PASSWORD)
                server.send_message(msg)
                print(f"✅ Email (Part {part_num}) 發送成功！")
                return True
                
        except Exception as e:
            print(f"❌ Email (Part {part_num}) 發送失敗: {e}")
            return False

    def pack_and_send_simple(self, excel_report_path: str):
        """
        簡單版本：打包所有內容到單一 ZIP 檔案並發送
        如果檔案太大會警告但仍嘗試發送，有可能報錯失敗
        """
        if not self.valid:
            print("⚠️ EmailReporter 憑證無效，跳過郵寄。")
            return False
            
        print("📦 開始執行打包與郵寄...")
        
        # zip 檔名
        timestamp = datetime.now().strftime('%Y%m%d_%H%M')
        zip_filename = f"website_check_results_{timestamp}.zip"
        
        try:
            files_in_zip = []
            with zipfile.ZipFile(zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
                
                # 加入 Excel 報告
                if os.path.exists(excel_report_path):
                    print(f"    + 加入報告: {excel_report_path}")
                    zipf.write(excel_report_path, os.path.basename(excel_report_path))
                    files_in_zip.append(os.path.basename(excel_report_path))
                else:
                    print(f"⚠️ 找不到報告檔: {excel_report_path}")
                
                # 加入爬蟲執行日誌
                vm_log_path = os.path.expanduser('~/crawler_execution.log')
                if os.path.exists(vm_log_path):
                    print(f"    + 加入日誌: {vm_log_path}")
                    zipf.write(vm_log_path, "crawler_execution.log")
                    files_in_zip.append("crawler_execution.log")
                else:
                    print(f"⚠️ 找不到日誌檔: {vm_log_path}")

                # 加入 assets 資料夾
                if os.path.exists("assets"):
                    print("    + 加入 assets 資料夾...")
                    for root, dirs, files in os.walk("assets"):
                        for file in files:
                            # 跳過 HTML 檔案
                            # if file.lower().endswith('.html'):
                            #     continue
                            
                            file_path = os.path.join(root, file)
                            arcname = os.path.relpath(file_path, start=".")
                            zipf.write(file_path, arcname)
                else:
                    print("⚠️ 找不到 assets 資料夾")
            
            print(f"✅ 壓縮完成: {zip_filename}")

            # 檢查檔案大小
            zip_size_mb = os.path.getsize(zip_filename) / 1024 / 1024
            if zip_size_mb > 25:
                print(f"⚠️ 警告：檔案大小 {zip_size_mb:.2f} MB 超過 Gmail 25MB 限制")
                print("💡 建議使用 pack_and_send_multi_part() 方法進行分割發送")

            # 發送郵件
            success = self._send_part(zip_filename, 1, 1, files_in_zip)
            
            # 清理檔案
            if os.path.exists(zip_filename):
                os.remove(zip_filename)
                print(f"🗑️ 已清理暫存檔案: {zip_filename}")
            
            return success

        except Exception as e:
            print(f"❌ 打包或發送過程發生錯誤: {e}")
            # 清理可能存在的暫存檔案
            if os.path.exists(zip_filename):
                os.remove(zip_filename)
            return False

    def pack_and_send_seperate(self, excel_report_path: str):
        """
        分割打包，避免超過 Gmail 大小限制
        每個網站資料夾分別打包，確保不超過 20MB
        """
        if not self.valid:
            print("⚠️ EmailReporter 憑證無效，跳過郵寄。")
            return False
            
        print("📦 開始執行分割打包與郵寄...")
        
        zip_files_info = []  # (檔案名, 檔案大小, 內容清單)
        max_size_mb = 20 # 每個包的最大大小 (MB)
        max_size_bytes = max_size_mb * 1024 * 1024
        
        # === 第一包：(Excel + Log) ===
        print("準備(Excel + Log)...")
        first_zip_filename = "crawl_data_part1.zip"
        first_files_list = []
        
        with zipfile.ZipFile(first_zip_filename, 'w', zipfile.ZIP_DEFLATED) as zipf:
            # 加入 Excel 報告
            if os.path.exists(excel_report_path):
                print(f"加入報告: {excel_report_path}")
                zipf.write(excel_report_path, os.path.basename(excel_report_path))
                first_files_list.append(os.path.basename(excel_report_path))
            else:
                print(f"⚠️ 找不到 Excel 報告: {excel_report_path}")
            
            # 加入爬蟲執行日誌
            vm_log_path = os.path.expanduser('~/crawler_execution.log')
            if os.path.exists(vm_log_path):
                print(f"加入日誌: {vm_log_path}")
                zipf.write(vm_log_path, "crawler_execution.log")
                first_files_list.append("crawler_execution.log")
            else:
                print(f"⚠️ 找不到執行日誌: {vm_log_path}")
        
        # 記錄第一個 ZIP 檔案資訊
        first_size = os.path.getsize(first_zip_filename)
        zip_files_info.append((first_zip_filename, first_size, first_files_list))
        first_size_mb = first_size//1024//1024
        print(f"Part 1 完成 ({first_size_mb:.1f} MB, 基本檔案)")
        
        # === 後續包：按大小動態處理 assets 資料夾 ===
        assets_dir = "assets"
        website_folders = []
        if os.path.exists(assets_dir):
            website_folders = [d for d in os.listdir(assets_dir) if os.path.isdir(os.path.join(assets_dir, d))]
            print(f"找到 {len(website_folders)} 個網站資料夾")
        else:
            print("⚠️ 沒有找到 assets 資料夾")

        if website_folders:
            part_num = 2
            website_index = 0
            
            while website_index < len(website_folders):
                current_zip_filename = f"crawl_data_part{part_num}.zip"
                files_in_current_zip = []
                
                print(f"  準備 Part {part_num} (網站資料)...")
                
                zipf = zipfile.ZipFile(current_zip_filename, 'w', zipfile.ZIP_DEFLATED)
                
                try:
                    # 逐一加入網站資料夾，檢查大小限制
                    while website_index < len(website_folders):
                        folder_name = website_folders[website_index]
                        folder_path = os.path.join(assets_dir, folder_name)
                        
                        # 先加入這個資料夾的所有檔案
                        for root, dirs, files in os.walk(folder_path):
                            for file in files:
                                # 跳過 HTML 檔案
                                # if file.lower().endswith('.html'):
                                #     continue
                                
                                file_path = os.path.join(root, file)
                                arcname = os.path.relpath(file_path, start=".")
                                zipf.write(file_path, arcname)
                        
                        files_in_current_zip.append(f"{folder_name}/ (網站資料夾)")
                        website_index += 1
                        
                        # 檢查當前大小
                        zipf.close()
                        current_size = os.path.getsize(current_zip_filename)
                        
                        # 如果超過限制且不是第一個資料夾，就停止加入更多
                        if current_size > max_size_bytes and len(files_in_current_zip) > 1:
                            print(f"⚠️ 達到 {current_size//1024//1024}MB 限制，此包完成")
                            break
                        elif website_index < len(website_folders):
                            # 還沒超過限制，重新開啟準備加入下一個
                            zipf = zipfile.ZipFile(current_zip_filename, 'a', zipfile.ZIP_DEFLATED)
                    
                finally:
                    if zipf.fp is not None and not zipf.fp.closed:
                        zipf.close()
                
                # 記錄 ZIP 檔案資訊
                final_size = os.path.getsize(current_zip_filename)
                zip_files_info.append((current_zip_filename, final_size, files_in_current_zip))
                final_size_mb = final_size//1024//1024
                print(f"Part {part_num} 完成 ({final_size_mb:.1f} MB, 包含 {len(files_in_current_zip)} 個網站)")
                part_num += 1
        
        total_parts = len(zip_files_info)
        print(f"📦 所有 ZIP 包準備完成！總共 {total_parts} 個包")
        
        # === 第二階段：依序寄送所有 ZIP 檔案 ===
        print("\n📧 開始寄送所有包...")
        success_count = 0
        
        for i, (zip_filename, zip_size, file_list) in enumerate(zip_files_info, 1):
            zip_size_mb = zip_size / 1024 / 1024
            print(f"  寄送 Part {i}/{total_parts} ({zip_size_mb:.1f} MB)...")
            
            if self._send_part(zip_filename, i, total_parts, file_list):
                success_count += 1
                print(f"    ✅ Part {i} 寄送成功")
            else:
                print(f"    ❌ Part {i} 寄送失敗")
            
            # 清理已寄送的檔案
            os.remove(zip_filename)
        
        print(f"\n🎉 多批次郵寄完成！成功發送 {success_count}/{total_parts} 個包")
        return success_count == total_parts