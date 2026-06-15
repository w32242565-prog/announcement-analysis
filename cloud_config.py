"""
云端/本地 统一配置模块

通过环境变量切换数据源，无需修改业务代码。

核心环境变量：
    DATABASE_URL        数据库路径或URL（默认: gem_stock.db）
                        支持 sqlite:///path, postgresql://, mysql://
    DATA_BASE_URL       原始数据（CSV等）的基础路径或URL前缀
                        本地示例: ./ 或空
                        云端示例: https://my-bucket.oss-cn-beijing.aliyuncs.com/gem_data/
    PDF_CACHE_DIR       PDF缓存目录（默认: 系统临时目录/pdf_cache）
    OUTPUT_DIR          图表/报告输出目录（默认: 当前目录）
    REPORT_BASE_URL     报告下载的基础URL（用于前端展示）

SQLite云端同步（可选）：
    CLOUD_DB_DOWNLOAD_URL   启动时从该URL下载数据库到本地
    CLOUD_DB_UPLOAD_URL     退出时上传数据库到该URL（需配合 sync_db_to_cloud）
    CLOUD_DB_API_KEY        上传时使用的API Key / Token

使用示例：
    from cloud_config import get_db_connection, read_csv_from_source, OUTPUT_DIR
    conn = get_db_connection()
    df = read_csv_from_source('01_公司基本信息/01_公司概况.csv', encoding='utf-8-sig')
"""

import os
import sys
import tempfile
import requests
import pandas as pd
import sqlite3

# 兼容 Windows 终端 UTF-8 输出
if hasattr(sys.stdout, 'reconfigure'):
    sys.stdout.reconfigure(encoding='utf-8')

# ============================
# 1. 基础配置（从环境变量读取，默认保持本地行为）
# ============================
USE_CLOUD = os.getenv("USE_CLOUD", "false").lower() in ("true", "1", "yes")

# 数据基础路径：本地目录或云端 URL 前缀
# 示例本地: "" 或 "./"
# 示例云端: "https://my-bucket.oss-cn-beijing.aliyuncs.com/gem_data"
DATA_BASE_URL = os.getenv("DATA_BASE_URL", "").rstrip("/")

# 数据库连接字符串
# SQLite本地文件: "gem_stock.db"
# SQLite云路径:   "sqlite:///mnt/data/gem_stock.db"
# PostgreSQL:     "postgresql://user:pass@host:port/db"
# MySQL:          "mysql+pymysql://user:pass@host:port/db"
DATABASE_URL = os.getenv("DATABASE_URL", "gem_stock.db")

# PDF 缓存目录（默认系统临时目录，云端实例通常只有 /tmp 可写）
PDF_CACHE_DIR = os.getenv("PDF_CACHE_DIR", os.path.join(tempfile.gettempdir(), "pdf_cache"))

# 图表/报告输出目录
OUTPUT_DIR = os.getenv("OUTPUT_DIR", ".")

# 报告下载基础 URL（前端展示用，如 https://cdn.example.com/reports/）
REPORT_BASE_URL = os.getenv("REPORT_BASE_URL", "").rstrip("/")

# SQLite 云端同步配置
CLOUD_DB_DOWNLOAD_URL = os.getenv("CLOUD_DB_DOWNLOAD_URL", "")
CLOUD_DB_UPLOAD_URL = os.getenv("CLOUD_DB_UPLOAD_URL", "")
CLOUD_DB_API_KEY = os.getenv("CLOUD_DB_API_KEY", "")


# ============================
# 2. 路径解析
# ============================
def resolve_path(local_path: str) -> str:
    """
    将本地相对路径解析为实际路径。
    若配置了 DATA_BASE_URL，则拼接为完整 URL 或路径。
    """
    if DATA_BASE_URL:
        return f"{DATA_BASE_URL}/{local_path.replace(os.sep, '/')}"
    return local_path


def resolve_output_path(filename: str) -> str:
    """解析输出文件路径：优先 OUTPUT_DIR"""
    if os.path.isabs(filename):
        return filename
    return os.path.join(OUTPUT_DIR, filename)


def ensure_dir(path: str):
    """确保目录存在"""
    dir_path = os.path.dirname(path)
    if dir_path:
        os.makedirs(dir_path, exist_ok=True)


# ============================
# 3. CSV 读取（支持本地 + HTTP(S) 云端）
# ============================
def read_csv_from_source(local_path: str, **pandas_kwargs) -> pd.DataFrame:
    """
    读取 CSV：支持本地文件或 HTTP(S) 云端 URL。
    如果是云端 URL，先下载到临时文件再读取。
    """
    path = resolve_path(local_path)

    if path.startswith(("http://", "https://")):
        resp = requests.get(path, timeout=60)
        resp.raise_for_status()

        suffix = os.path.splitext(local_path)[1] or ".csv"
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix, mode="wb") as f:
            f.write(resp.content)
            tmp_path = f.name

        try:
            df = pd.read_csv(tmp_path, **pandas_kwargs)
        finally:
            os.unlink(tmp_path)
        return df
    else:
        return pd.read_csv(path, **pandas_kwargs)


# ============================
# 4. 数据库连接（优先 SQLite，可扩展）
# ============================
def get_db_path() -> str:
    """获取 SQLite 数据库文件路径"""
    if DATABASE_URL.startswith("sqlite:///"):
        return DATABASE_URL.replace("sqlite:///", "")
    return DATABASE_URL


def _sync_db_from_cloud(url: str, local_path: str):
    """从云端同步数据库文件到本地"""
    ensure_dir(local_path)
    print(f"[cloud_config] 正在从云端同步数据库: {url} -> {local_path}")
    resp = requests.get(url, timeout=120)
    resp.raise_for_status()
    with open(local_path, "wb") as f:
        f.write(resp.content)
    print(f"[cloud_config] 数据库同步完成: {local_path}")


def get_db_connection():
    """
    获取数据库连接。
    当前主要支持 SQLite；如需 PostgreSQL/MySQL，可安装 sqlalchemy 扩展。
    """
    db_path = get_db_path()

    # 如果配置了云端下载地址，先同步到本地
    if CLOUD_DB_DOWNLOAD_URL and USE_CLOUD:
        _sync_db_from_cloud(CLOUD_DB_DOWNLOAD_URL, db_path)

    conn = sqlite3.connect(db_path)
    return conn


def sync_db_to_cloud(local_path: str = None, upload_url: str = None):
    """
    将本地数据库同步到云端。
    支持 HTTP PUT / POST 上传，或通过环境变量配置。
    """
    local_path = local_path or get_db_path()
    upload_url = upload_url or CLOUD_DB_UPLOAD_URL

    if not upload_url or not os.path.exists(local_path):
        return False

    headers = {}
    if CLOUD_DB_API_KEY:
        headers["Authorization"] = f"Bearer {CLOUD_DB_API_KEY}"

    with open(local_path, "rb") as f:
        # 优先尝试 PUT
        resp = requests.put(upload_url, data=f, headers=headers, timeout=120)
        if resp.status_code not in (200, 201, 204):
            # 回退 POST
            f.seek(0)
            resp = requests.post(upload_url, files={"file": f}, headers=headers, timeout=120)

    if resp.status_code in (200, 201, 204):
        print(f"[cloud_config] 数据库已同步到云端")
        return True
    else:
        print(f"[cloud_config] 数据库同步失败: HTTP {resp.status_code}")
        return False


# ============================
# 5. 便捷工具
# ============================
def get_pdf_cache_dir() -> str:
    """获取 PDF 缓存目录，并确保存在"""
    os.makedirs(PDF_CACHE_DIR, exist_ok=True)
    return PDF_CACHE_DIR


def print_config():
    """打印当前配置（调试用）"""
    print("=" * 60)
    print("cloud_config 当前配置")
    print("=" * 60)
    print(f"  USE_CLOUD            = {USE_CLOUD}")
    print(f"  DATA_BASE_URL        = {DATA_BASE_URL or '(未配置，使用本地路径)'}")
    print(f"  DATABASE_URL         = {DATABASE_URL}")
    print(f"  PDF_CACHE_DIR        = {PDF_CACHE_DIR}")
    print(f"  OUTPUT_DIR           = {OUTPUT_DIR}")
    print(f"  REPORT_BASE_URL      = {REPORT_BASE_URL or '(未配置)'}")
    print(f"  CLOUD_DB_DOWNLOAD_URL= {CLOUD_DB_DOWNLOAD_URL or '(未配置)'}")
    print(f"  CLOUD_DB_UPLOAD_URL  = {CLOUD_DB_UPLOAD_URL or '(未配置)'}")
    print("=" * 60)


if __name__ == "__main__":
    print_config()
