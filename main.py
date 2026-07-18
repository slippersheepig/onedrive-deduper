import os
import sys
import json
import requests
import msal

# 配置参数 (通过环境变量传入)
CLIENT_ID = os.getenv("CLIENT_ID")
SHAREPOINT_HOST = os.getenv("SHAREPOINT_HOST") # 例如: contoso.sharepoint.com
SHAREPOINT_SITE = os.getenv("SHAREPOINT_SITE", "") # 例如: /sites/CommunicationSite
TARGET_FOLDER = os.getenv("TARGET_FOLDER", "/照片")
DRY_RUN = os.getenv("DRY_RUN", "true").lower() == "true"
TG_BOT_TOKEN = os.getenv("TG_BOT_TOKEN", "")
TG_CHAT_ID = os.getenv("TG_CHAT_ID", "")

# 固定的 MSAL 配置
AUTHORITY = "https://login.microsoftonline.com/common"
SCOPES = ["Files.ReadWrite.All", "offline_access"]
CACHE_FILE = "/app/data/token_cache.json"

def send_tg_msg(text):
    if TG_BOT_TOKEN and TG_CHAT_ID:
        url = f"https://api.telegram.org/bot{TG_BOT_TOKEN}/sendMessage"
        requests.post(url, json={"chat_id": TG_CHAT_ID, "text": text})
        print(f"Telegram 推送: {text}")

def load_cache():
    cache = msal.SerializableTokenCache()
    if os.path.exists(CACHE_FILE):
        with open(CACHE_FILE, "r") as f:
            cache.deserialize(f.read())
    return cache

def save_cache(cache):
    if cache.has_state_changed:
        with open(CACHE_FILE, "w") as f:
            f.write(cache.serialize())

def get_access_token():
    cache = load_cache()
    app = msal.PublicClientApplication(CLIENT_ID, authority=AUTHORITY, token_cache=cache)
    
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            save_cache(cache)
            return result["access_token"]

    flow = app.initiate_device_flow(scopes=SCOPES)
    if "user_code" not in flow:
        raise Exception("无法初始化设备流认证。")
    
    print("=" * 50)
    print(flow["message"])
    print("=" * 50)
    sys.stdout.flush()
    
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" in result:
        save_cache(cache)
        return result["access_token"]
    else:
        raise Exception(f"认证失败: {result.get('error_description')}")

def get_drive_and_folder(access_token):
    headers = {"Authorization": f"Bearer {access_token}"}
    
    # 动态构建基于 SharePoint 或个人 OneDrive 的 Graph API 端点
    if SHAREPOINT_HOST:
        site_path = f":{SHAREPOINT_SITE}:" if SHAREPOINT_SITE and not SHAREPOINT_SITE.startswith(":") else SHAREPOINT_SITE
        if site_path and not site_path.endswith(":"):
            site_path += ":"
        if site_path == "::" or not site_path: 
            site_path = ""
            
        base_url = f"https://graph.microsoft.com/v1.0/sites/{SHAREPOINT_HOST}{site_path}/drive"
    else:
        base_url = "https://graph.microsoft.com/v1.0/me/drive"
        
    # 获取 drive_id
    drive_res = requests.get(base_url, headers=headers).json()
    if "error" in drive_res:
        raise Exception(f"获取 Drive 失败: {drive_res['error']['message']}")
    drive_id = drive_res["id"]
    
    # 获取 folder_id
    folder_url = f"{base_url}/root:{TARGET_FOLDER}"
    folder_res = requests.get(folder_url, headers=headers).json()
    if "error" in folder_res:
        raise Exception(f"获取目标文件夹失败: {folder_res['error']['message']}")
    folder_id = folder_res["id"]
    
    return drive_id, folder_id

def scan_and_dedupe(access_token, drive_id, folder_id):
    headers = {"Authorization": f"Bearer {access_token}"}
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{folder_id}/children"
    
    hash_dict = {}
    deleted_count = 0
    saved_size = 0
    
    while url:
        res = requests.get(url, headers=headers).json()
        if "value" not in res:
            break
            
        for item in res["value"]:
            if "file" in item:
                hashes = item["file"].get("hashes", {})
                file_hash = hashes.get("sha1Hash") or hashes.get("quickXorHash")
                file_size = item.get("size", 0)
                item_id = item["id"]
                item_name = item["name"]
                
                if not file_hash:
                    continue
                
                dict_key = f"{file_hash}_{file_size}"
                
                if dict_key in hash_dict:
                    deleted_count += 1
                    saved_size += file_size
                    print(f"[发现重复] {item_name} (原文件: {hash_dict[dict_key]})")
                    
                    if not DRY_RUN:
                        del_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}"
                        requests.delete(del_url, headers=headers)
                        print(f" -> 已移入回收站")
                else:
                    hash_dict[dict_key] = item_name

        url = res.get("@odata.nextLink")
        
    return deleted_count, saved_size

def main():
    if not CLIENT_ID:
        print("缺少环境变量: CLIENT_ID")
        return

    print(f"启动扫描... (DRY_RUN={DRY_RUN})")
    try:
        token = get_access_token()
        drive_id, folder_id = get_drive_and_folder(token)
        count, size = scan_and_dedupe(token, drive_id, folder_id)
        
        size_mb = size / (1024 * 1024)
        msg = f"SharePoint/OneDrive 清理完成。
状态: {'空跑测试' if DRY_RUN else '真实执行'}
发现并清理: {count} 个文件
释放空间: {size_mb:.2f} MB"
        print(msg)
        if count > 0:
            send_tg_msg(msg)
            
    except Exception as e:
        err_msg = f"去重脚本执行异常: {str(e)}"
        print(err_msg)
        send_tg_msg(err_msg)

if __name__ == "__main__":
    main()
