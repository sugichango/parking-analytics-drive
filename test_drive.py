import streamlit as st
import json
from google.oauth2 import service_account
from googleapiclient.discovery import build

def test_google_drive_broad():
    print("--- Google Drive 総当たり検索テスト ---")
    
    try:
        # 1. 認証情報の作成
        info = json.loads(st.secrets["gcp_service_account"])
        email = info.get('client_email')
        print(f"✅ 使用中の合鍵: {email}")
        
        creds = service_account.Credentials.from_service_account_info(info)
        service = build('drive', 'v3', credentials=creds)

        # 2. 全ファイル一覧の取得テスト（フォルダ指定なし！）
        print(f"🔍 この合鍵がアクセスを許可されているファイルをすべて探しています...")
        # 共有ドライブなども含めて全検索
        results = service.files().list(
            pageSize=10, 
            fields="files(id, name, parents)",
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        files = results.get('files', [])

        if not files:
            print(f"❌ 警告: この合鍵（{email}）からは、ファイルが一つも見えません。")
            print("--- 対処法 ---")
            print("1. Googleドライブでフォルダを共有した相手のメールアドレスが、上記のアドレスと1文字も違わないか、今一度確認してください。")
            print("2. 共有の際、右下の「送信」ボタンを確実に押したか確認してください。")
        else:
            print(f"🎯 成功！以下のファイルへのアクセス権を確認しました:")
            for f in files:
                parents = f.get('parents', ['なし'])
                print(f"   - ファイル名: {f['name']}")
                print(f"     ID: {f['id']}")
                print(f"     親フォルダID: {parents[0]}")
                
        # フォルダIDの直接検証
        target_folder = st.secrets.get("google_drive_folder_id", "")
        print(f"\n📏 ターゲットのフォルダID検証: {target_folder}")
        try:
            folder_info = service.files().get(fileId=target_folder, fields="id, name, permissions", supportsAllDrives=True).execute()
            print(f"   ✅ フォルダ自体は見つかりました！ 名前: {folder_info.get('name')}")
        except Exception as e:
            print(f"   ❌ フォルダIDに直接アクセスできませんでした。権限がないか、IDが間違っている可能性があります。")

    except Exception as e:
        print(f"❌ 致命的なエラーが発生しました:\n{str(e)}")

if __name__ == "__main__":
    test_google_drive_broad()
