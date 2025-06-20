import os
import getpass

# Google Calendar API の認証スコープ
SCOPES = ["https://www.googleapis.com/auth/calendar"]

# 各ユーザーごとにトークンを保存するパス
USER_ID = getpass.getuser()
TOKEN_DIR = "tokens"
TOKEN_PATH = os.path.join(TOKEN_DIR, f"token_{USER_ID}.pickle")
