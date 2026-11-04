"""
database.py — Database Models, Migrations & Auto-Seeding for DailyShowroomsSales App
Supports SQLite (local standalone) & PostgreSQL / Supabase (Cloud First).
"""

import os
import sys
from datetime import datetime
from typing import Optional, List, Dict, Any

try:
    from dotenv import load_dotenv
    _env_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), ".env")
    if os.path.exists(_env_path):
        load_dotenv(dotenv_path=_env_path, override=True)
    else:
        _root_env = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".env")
        if os.path.exists(_root_env):
            load_dotenv(dotenv_path=_root_env, override=True)
except Exception:
    pass

from sqlalchemy import (
    create_engine, Column, Integer, String, Float, Boolean,
    DateTime, ForeignKey, LargeBinary, Index, text, event
)
from sqlalchemy.orm import declarative_base, sessionmaker, relationship, scoped_session

Base = declarative_base()

# =============================================================================
# MODELS
# =============================================================================

class Branch(Base):
    __tablename__ = "tbl_branches"

    BranchNo = Column(String(10), primary_key=True)  # e.g., '01', '02', '03', '04'
    BranchName = Column(String(100), nullable=False)

    cities = relationship("City", back_populates="branch", cascade="all, delete-orphan")

    def __repr__(self):
        return f"<Branch {self.BranchNo}: {self.BranchName}>"


class City(Base):
    __tablename__ = "tbl_citys"

    CityNo = Column(String(10), primary_key=True)  # e.g., '001', '002', '100', '200'
    BranchNo = Column(String(10), ForeignKey("tbl_branches.BranchNo"), nullable=False)
    CityName = Column(String(100), nullable=False)

    branch = relationship("Branch", back_populates="cities")
    showrooms = relationship("Showroom", back_populates="city")

    def __repr__(self):
        return f"<City {self.CityNo}: {self.CityName}>"


class Showroom(Base):
    __tablename__ = "tbl_showrooms"

    No = Column(String(20), primary_key=True)  # e.g., '101', '102', 'JAN'
    DESC = Column(String(255), nullable=False, index=True)  # Name from Odoo
    Short = Column(String(100), nullable=False, index=True) # Normalized display name
    CityNo = Column(String(10), ForeignKey("tbl_citys.CityNo"), nullable=True)
    is_active = Column(Boolean, default=True)

    city = relationship("City", back_populates="showrooms")

    def __repr__(self):
        return f"<Showroom {self.No}: {self.Short} ({self.DESC})>"


class OldClosing(Base):
    __tablename__ = "tbl_old"

    id = Column(Integer, primary_key=True, autoincrement=True)
    closing_no = Column(String(50), unique=True, nullable=False, index=True)  # e.g., 'POS/57861'
    showroom_short = Column(String(100), nullable=True, index=True)
    closing_date = Column(String(20), nullable=True)  # 'YYYY-MM-DD'
    net_total = Column(Float, nullable=True, default=0.0)
    batch_id = Column(String(50), ForeignKey("report_batches.batch_id"), nullable=True, index=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    batch = relationship("ReportBatch", back_populates="closings")

    def __repr__(self):
        return f"<OldClosing {self.closing_no} - {self.showroom_short}>"


class ReportBatch(Base):
    __tablename__ = "report_batches"

    batch_id = Column(String(50), primary_key=True)  # e.g. BATCH-20261027-142000
    report_date = Column(String(20), nullable=False, index=True)  # 'YYYY-MM-DD'
    sessions_count = Column(Integer, default=0)
    unique_showrooms_count = Column(Integer, default=0)
    total_sales = Column(Float, default=0.0)
    net_sales = Column(Float, default=0.0)
    is_rolled_back = Column(Boolean, default=False)
    
    # PDF Blobs stored for cloud backup & immediate download
    detailed_pdf_blob = Column(LargeBinary, nullable=True)
    totals_pdf_blob = Column(LargeBinary, nullable=True)
    color_pdf_blob = Column(LargeBinary, nullable=True)
    
    notes = Column(String(500), nullable=True)
    created_at = Column(DateTime, default=datetime.utcnow)

    closings = relationship("OldClosing", back_populates="batch")

    def __repr__(self):
        return f"<ReportBatch {self.batch_id} for {self.report_date}>"


# =============================================================================
# SEED DATA
# =============================================================================

BRANCHES_SEED = [
    {"BranchNo": "01", "BranchName": "الغربية"},
    {"BranchNo": "02", "BranchName": "الوسطى"},
    {"BranchNo": "03", "BranchName": "الشرقية"},
    {"BranchNo": "04", "BranchName": "الجنوبية"},
]

CITIES_SEED = [
    {"CityNo": "001", "BranchNo": "01", "CityName": "جدة"},
    {"CityNo": "002", "BranchNo": "01", "CityName": "مكة"},
    {"CityNo": "003", "BranchNo": "01", "CityName": "الطائف"},
    {"CityNo": "004", "BranchNo": "01", "CityName": "ينبع"},
    {"CityNo": "005", "BranchNo": "01", "CityName": "المدينة"},
    {"CityNo": "100", "BranchNo": "02", "CityName": "الرياض"},
    {"CityNo": "200", "BranchNo": "03", "CityName": "الدمام"},
    {"CityNo": "201", "BranchNo": "03", "CityName": "الاحساء"},
    {"CityNo": "202", "BranchNo": "03", "CityName": "الهفوف"},
    {"CityNo": "300", "BranchNo": "04", "CityName": "أبها"},
    {"CityNo": "301", "BranchNo": "04", "CityName": "خميس مشيط"},
    {"CityNo": "302", "BranchNo": "04", "CityName": "جيزان"},
    {"CityNo": "303", "BranchNo": "04", "CityName": "نجران"},
    {"CityNo": "304", "BranchNo": "04", "CityName": "محايل عسير"},
    {"CityNo": "305", "BranchNo": "04", "CityName": "أبو عريش"},
    {"CityNo": "306", "BranchNo": "04", "CityName": "صبيا"},
]

SHOWROOMS_SEED = [
    {"No": "001", "DESC": "001-JED Office", "Short": "001-JED Office", "CityNo": "001"},
    {"No": "007", "DESC": "مستودع التوالف", "Short": "مستودع التوالف", "CityNo": "001"},
    {"No": "008", "DESC": "مستودع التسائر", "Short": "مستودع التسائر", "CityNo": "001"},
    {"No": "009", "DESC": "مستودع الهدايا", "Short": "مستودع الهدايا", "CityNo": "001"},
    {"No": "101", "DESC": "معرض الشربتلي (Al-sharbatly Shrm)", "Short": "الشربتلي", "CityNo": "001"},
    {"No": "102", "DESC": "معرض البطاطي (Al-Batati - Shrm)", "Short": "البطاطي", "CityNo": "001"},
    {"No": "103", "DESC": "معرض الاسنوي (Al-Asnawyi-Shrm)", "Short": "الاسنوي", "CityNo": "001"},
    {"No": "104", "DESC": "معرض الذهب (Al-Dahab-Shrm)", "Short": "الذهب", "CityNo": "001"},
    {"No": "105", "DESC": "معرض باناجة (-BanagaShrm)", "Short": "باناجة", "CityNo": "001"},
    {"No": "106", "DESC": "معرض الصفاء ((-ALSafa Shrm", "Short": "الصفاء", "CityNo": "001"},
    {"No": "107", "DESC": "معرض الاوقاف (--.AlAwqafshrm)", "Short": "الاوقاف", "CityNo": "001"},
    {"No": "108", "DESC": "معرض العيدروس (.Al-Eidaros-Shrm)", "Short": "العيدروس", "CityNo": "001"},
    {"No": "109", "DESC": "معرض جرند اويسس مول", "Short": "جرند اويسس مول", "CityNo": "001"},
    {"No": "110", "DESC": "معرض الشجره اويسس مول", "Short": "الشجره اويسس مول", "CityNo": "001"},
    {"No": "111", "DESC": "معرض المكرونة", "Short": "المكرونة", "CityNo": "001"},
    {"No": "112", "DESC": "معرض البوادي. (Al-Bawadi-shrm)", "Short": "البوادي", "CityNo": "001"},
    {"No": "113", "DESC": "معرض الجامعة", "Short": "الجامعة", "CityNo": "001"},
    {"No": "114", "DESC": "معرض الكندرة", "Short": "الكندرة", "CityNo": "001"},
    {"No": "115", "DESC": "معرض مكة الستين (.Makah Alseten-shrm)", "Short": "الستين مكة", "CityNo": "002"},
    {"No": "116", "DESC": "معرض مكة الشرائع (Maka Alsharia-shrm)", "Short": "الشرائع مكة", "CityNo": "002"},
    {"No": "117", "DESC": "معرض مكة العتيبية (Maka Al-Otaibia-Shrm )", "Short": "العتيبية مكة", "CityNo": "002"},
    {"No": "118", "DESC": "معرض الطائف (--.AlTaifShrm)", "Short": "الطائف", "CityNo": "003"},
    {"No": "119", "DESC": "معرض ينبع (.Yanbu Shrm)", "Short": "ينبع", "CityNo": "004"},
    {"No": "120", "DESC": "معرض الحمدانية (Al-Hamdaniya-shrm)", "Short": "الحمدانية", "CityNo": "001"},
    {"No": "121", "DESC": "معرض الدراجة (Al-Daragah ms)", "Short": "الدراجة", "CityNo": "001"},
    {"No": "122", "DESC": "معرض البوادي - MS- (Al-Bawadi)", "Short": "البوادي MS", "CityNo": "001"},
    {"No": "123", "DESC": "معرض قربان (qurban ms-shrm)", "Short": "قربان", "CityNo": "005"},
    {"No": "124", "DESC": "معرض الخاسكية MS", "Short": "الخاسكية MS", "CityNo": "001"},
    {"No": "125", "DESC": "معرض مكة العزيزية (Maka Al-Azezia-Shrm)", "Short": "العزيزة مكة", "CityNo": "002"},
    {"No": "126", "DESC": "معرض السلام مول", "Short": "السلام", "CityNo": "001"},
    {"No": "127", "DESC": "معرض الوزيرية (Al-Waziriyah-Shrm)", "Short": "الوزيرية", "CityNo": "001"},
    {"No": "128", "DESC": "معرض الشاطئ (Al-Shata-Shrm)", "Short": "الشاطئ", "CityNo": "001"},
    {"No": "129", "DESC": "معرض المكرونة 2 (Al-Makarona 2-shrm)", "Short": "المكرونة ٢", "CityNo": "001"},
    {"No": "130", "DESC": "معرض الخزان. (Al-khazan-shrm)", "Short": "الخزان", "CityNo": "100"},
    {"No": "131", "DESC": "معرض البطحاء (--.Albathashrm)", "Short": "البطحاء", "CityNo": "100"},
    {"No": "132", "DESC": "معرض المعيقلية (Al-Muaqiliah)", "Short": "المعيقلية", "CityNo": "100"},
    {"No": "133", "DESC": "معرض النسيم (Al-Nasem ms-shrm)", "Short": "النسيم", "CityNo": "100"},
    {"No": "134", "DESC": "معرض العروبة (Al-Auruba-Shrm)", "Short": "العروبة", "CityNo": "100"},
    {"No": "140", "DESC": "معرض الدمام (.Al-Dammam-shrm)", "Short": "الدمام", "CityNo": "200"},
    {"No": "141", "DESC": "معرض الاحساء", "Short": "الاحساء", "CityNo": "201"},
    {"No": "142", "DESC": "معرض الاحساء الهفوف (Al-Ihsaa Alhafof shrm)", "Short": "الاحساء الهفوف", "CityNo": "202"},
    {"No": "143", "DESC": "معرض الحب (Alhob ms-shrm)", "Short": "الحب", "CityNo": "200"},
    {"No": "144", "DESC": "معرض الجملة MS", "Short": "الجملة", "CityNo": "200"},
    {"No": "145", "DESC": "معرض حفر الباطن (Hafar Al-Batin shrm)", "Short": "الباطن", "CityNo": None},
    {"No": "150", "DESC": "معرض أبهاء", "Short": "ابها", "CityNo": "300"},
    {"No": "151", "DESC": "معرض العسكرية (.AlAskareah--shrm)", "Short": "العسكرية", "CityNo": "301"},
    {"No": "152", "DESC": "معرض خميس البلد (.Khames Albalad-shrm)", "Short": "خميس البلد", "CityNo": "301"},
    {"No": "153", "DESC": "معرض محايل عسير (Mohayel Aseer-shrm)", "Short": "محايل عسير", "CityNo": "304"},
    {"No": "154", "DESC": "معرض جيزان الذهب (Gaizan Althahb-shrm )", "Short": "جيزان الذهب", "CityNo": "302"},
    {"No": "155", "DESC": "معرض جيزان الاسرة (Jaizan Alusrah-shrm)", "Short": "جيزان الاسرة", "CityNo": "302"},
    {"No": "156", "DESC": "معرض ابو عريش MS closingxx (Abo ARESH-SHRM)", "Short": "ابو عريش XXX", "CityNo": "305"},
    {"No": "157", "DESC": "معرض صبيا (-.Sabiashrm)", "Short": "صبيا", "CityNo": "306"},
    {"No": "158", "DESC": "معرض النورية (Nouriyah)", "Short": "النورية", "CityNo": "002"},
    {"No": "159", "DESC": "معرض ميس برتي (Miss Party)", "Short": "ميس برتي", "CityNo": "001"},
    {"No": "JAN", "DESC": "مستودع الجنوب", "Short": "مستودع الجنوب", "CityNo": "300"},
    {"No": "JEDHA", "DESC": "مستودع جدة الشعر", "Short": "مستودع جدة الشعر", "CityNo": "001"},
    {"No": "JEDMK", "DESC": "مستودع جدة المكياج", "Short": "مستودع جدة المكياج", "CityNo": "001"},
    {"No": "JEDPE", "DESC": "مستودع جدة العطور", "Short": "مستودع جدة العطور", "CityNo": "001"},
    {"No": "RIY", "DESC": "مستودع الرياض", "Short": "مستودع الرياض", "CityNo": "100"},
]


# =============================================================================
# DATABASE ENGINE & SESSIONS
# =============================================================================

import urllib.parse

def clean_pg_url(url: str) -> str:
    """Normalizes PostgreSQL URL, handles special characters in passwords, and ensures SSL."""
    url = (url or "").strip()
    if not url or "sqlite" in url.lower():
        return url
    if url.startswith("postgres://"):
        url = url.replace("postgres://", "postgresql+psycopg2://", 1)
    elif url.startswith("postgresql://") and not url.startswith("postgresql+"):
        url = url.replace("postgresql://", "postgresql+psycopg2://", 1)

    if "@" in url:
        parts = url.split("://", 1)
        proto = parts[0]
        rest = parts[1]
        last_at = rest.rfind("@")
        first_colon = rest.find(":")
        if first_colon != -1 and first_colon < last_at:
            username = rest[:first_colon]
            raw_password = rest[first_colon+1:last_at]
            host_and_beyond = rest[last_at+1:]
            encoded_password = urllib.parse.quote_plus(urllib.parse.unquote_plus(raw_password))
            url = f"{proto}://{username}:{encoded_password}@{host_and_beyond}"

    if "sslmode=" not in url:
        delimiter = "&" if "?" in url else "?"
        url = f"{url}{delimiter}sslmode=require"

    return url

DEFAULT_DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "daily_sales.db")
SUPABASE_CLOUD_URL = "postgresql+psycopg2://postgres.ztalhuqyeqdusxmuklmw:Abd%406326653@aws-0-ap-northeast-1.pooler.supabase.com:5432/postgres?sslmode=require"

# Support environment variables, .env, and Streamlit Secrets (st.secrets)
_raw_db_url = os.environ.get("DATABASE_URL")
if not _raw_db_url:
    try:
        import streamlit as st
        if hasattr(st, "secrets") and "DATABASE_URL" in st.secrets:
            _raw_db_url = st.secrets["DATABASE_URL"]
    except Exception:
        pass

# If deployed in cloud (Linux container), default directly to Supabase cloud database
if not _raw_db_url and sys.platform != "win32":
    _raw_db_url = SUPABASE_CLOUD_URL

DATABASE_URL = clean_pg_url(_raw_db_url or f"sqlite:///{DEFAULT_DB_PATH}")

is_cloud_db = not ("sqlite" in DATABASE_URL.lower())

engine_kwargs = {}
if is_cloud_db:
    # Cloud-resilient connection pool settings with auto-reconnect
    engine_kwargs = {
        "pool_size": 10,
        "max_overflow": 20,
        "pool_pre_ping": True,
        "pool_recycle": 300
    }
else:
    engine_kwargs = {
        "connect_args": {"check_same_thread": False}
    }

engine = create_engine(
    DATABASE_URL,
    echo=False,
    **engine_kwargs
)

@event.listens_for(engine, "connect")
def set_sqlite_pragma(dbapi_connection, connection_record):
    """Enables SQLite WAL Mode, fast sync, and in-memory cache for high-concurrency throughput."""
    if not is_cloud_db:
        try:
            cursor = dbapi_connection.cursor()
            cursor.execute("PRAGMA journal_mode=WAL")
            cursor.execute("PRAGMA synchronous=NORMAL")
            cursor.execute("PRAGMA cache_size=-64000")  # 64MB cache
            cursor.execute("PRAGMA temp_store=MEMORY")
            cursor.close()
        except Exception:
            pass

SessionLocal = scoped_session(sessionmaker(autocommit=False, autoflush=False, bind=engine))

def get_database_status_info() -> Dict[str, Any]:
    """Returns database connection type and status."""
    masked = "Local Standalone"
    if is_cloud_db:
        if "@" in DATABASE_URL:
            masked = DATABASE_URL.split("@")[-1].split("?")[0]
        else:
            masked = "Cloud PostgreSQL"
    return {
        "is_cloud": is_cloud_db,
        "type": "Cloud PostgreSQL (Supabase)" if is_cloud_db else "Local SQLite (Standalone)",
        "url_masked": masked
    }

def get_db():
    db = SessionLocal()
    try:
        yield db
    finally:
        db.close()

def init_db(db_session=None):
    """
    Initializes tables and seeds reference data if not already present.
    """
    Base.metadata.create_all(bind=engine)
    session = db_session or SessionLocal()
    try:
        # 1. Seed Branches
        if session.query(Branch).count() == 0:
            for b_data in BRANCHES_SEED:
                session.add(Branch(**b_data))
            session.commit()

        # 2. Seed Cities
        if session.query(City).count() == 0:
            for c_data in CITIES_SEED:
                session.add(City(**c_data))
            session.commit()

        # 3. Seed Showrooms
        if session.query(Showroom).count() == 0:
            for s_data in SHOWROOMS_SEED:
                session.add(Showroom(**s_data))
            session.commit()

        # 4. Auto-Seed Initial tbl_old from seed CSV if tbl_old is empty
        if session.query(OldClosing).count() == 0:
            seed_csv_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "test_data", "tbl_old_seed.csv")
            if os.path.exists(seed_csv_path):
                import csv
                objects_to_add = []
                seen_c_nos = set()
                with open(seed_csv_path, mode="r", encoding="utf-8-sig") as f:
                    reader = csv.DictReader(f)
                    for row in reader:
                        c_no = row.get("ClosingNO", "").strip()
                        if c_no and c_no not in seen_c_nos:
                            seen_c_nos.add(c_no)
                            net_val = 0.0
                            try:
                                net_val = float(row.get("NetTotal", 0.0) or 0.0)
                            except Exception:
                                pass
                            objects_to_add.append(OldClosing(
                                closing_no=c_no,
                                showroom_short=row.get("Showroom", "").strip(),
                                closing_date=row.get("ClosingDate", "").strip(),
                                net_total=net_val,
                                batch_id="INITIAL_SEED"
                            ))
                if objects_to_add:
                    session.bulk_save_objects(objects_to_add)
                session.commit()

    except Exception as e:
        session.rollback()
        raise e
    finally:
        if db_session is None:
            session.close()

if __name__ == "__main__":
    init_db()
    print("Database initialized and seeded successfully!")
