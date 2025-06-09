from pymongo import MongoClient
from datetime import datetime
import hashlib
import bcrypt
import json
import os
import logging

# 設定日誌
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# MongoDB 連接設定
MONGO_URI = "mongodb+srv://toma94046:gGIuVjyORUgqmlgz@cluster0.rg42vj9.mongodb.net/?retryWrites=true&w=majority&appName=Cluster0"
DB_NAME = "course_signin"
QR_SEED = "secure_seed_2024"

# 預設密碼設定
DEFAULT_ADMIN_PASSWORD = "admin123"
DEFAULT_USER_PASSWORD = "user123"

def load_settings():
    """載入設定檔"""
    try:
        if os.path.exists("settings.json"):
            with open("settings.json", "r", encoding="utf-8") as f:
                settings = json.load(f)
                return settings
    except Exception as e:
        logger.error(f"載入設定檔時發生錯誤: {str(e)}")
    return {}

def save_settings(settings):
    """儲存設定檔"""
    try:
        with open("settings.json", "w", encoding="utf-8") as f:
            json.dump(settings, f, ensure_ascii=False, indent=4)
    except Exception as e:
        logger.error(f"儲存設定檔時發生錯誤: {str(e)}")

# 建立全域連接池
_client = None

def get_client():
    """獲取 MongoDB 客戶端連接"""
    global _client
    if _client is None:
        try:
            _client = MongoClient(MONGO_URI)
            logger.info("MongoDB 連接成功")
        except Exception as e:
            logger.error(f"MongoDB 連接失敗: {str(e)}")
            raise
    return _client

def get_db():
    """獲取資料庫連接"""
    return get_client()[DB_NAME]

def init_db():
    """初始化資料庫結構"""
    try:
        logger.info("開始初始化資料庫...")
        db = get_db()
        
        # 創建集合（如果不存在）
        collections = ['org_info', 'classes', 'sessions', 'attendees', 'checkins', 'users']
        for collection in collections:
            if collection not in db.list_collection_names():
                db.create_collection(collection)
                logger.info(f"建立集合: {collection}")
        
        # 初始化組織資訊
        org_info = db.org_info.find_one({"_id": 1})
        if not org_info:
            db.org_info.insert_one({
                "_id": 1,
                "org_name": "課堂簽到系統",
                "manager": "",
                "contact": ""
            })
            logger.info("初始化組織資訊")
        
        # 初始化預設使用者帳號
        settings = load_settings()
        admin_password = settings.get("admin_password", DEFAULT_ADMIN_PASSWORD)
        user_password = settings.get("user_password", DEFAULT_USER_PASSWORD)
        
        # 建立管理員帳號
        admin = db.users.find_one({"username": "admin"})
        if not admin:
            hashed = bcrypt.hashpw(admin_password.encode('utf-8'), bcrypt.gensalt())
            db.users.insert_one({
                "username": "admin",
                "password": hashed,
                "role": "admin",
                "created_at": datetime.now()
            })
            logger.info("建立管理員帳號")
        
        # 建立預設一般使用者帳號
        default_user = db.users.find_one({"username": "user"})
        if not default_user:
            hashed = bcrypt.hashpw(user_password.encode('utf-8'), bcrypt.gensalt())
            db.users.insert_one({
                "username": "user",
                "password": hashed,
                "role": "user",
                "created_at": datetime.now()
            })
            logger.info("建立預設使用者帳號")
            
        logger.info("資料庫初始化完成")
        
    except Exception as e:
        logger.error(f"初始化資料庫時發生錯誤: {str(e)}")
        raise

def verify_user(username, password):
    """驗證使用者帳號密碼"""
    try:
        db = get_db()
        user = db.users.find_one({"username": username})
        if user and bcrypt.checkpw(password.encode('utf-8'), user['password']):
            return {
                "username": user['username'],
                "role": user['role']
            }
        return None
    except Exception as e:
        logger.error(f"驗證使用者時發生錯誤: {str(e)}")
        return None

def add_user(username, password, role="user"):
    """新增使用者"""
    try:
        db = get_db()
        if db.users.find_one({"username": username}):
            return False, "使用者名稱已存在"
        
        hashed = bcrypt.hashpw(password.encode('utf-8'), bcrypt.gensalt())
        db.users.insert_one({
            "username": username,
            "password": hashed,
            "role": role,
            "created_at": datetime.now()
        })
        logger.info(f"新增使用者: {username}")
        return True, "使用者新增成功"
    except Exception as e:
        logger.error(f"新增使用者時發生錯誤: {str(e)}")
        return False, f"新增使用者時發生錯誤: {str(e)}"

def change_password(username, old_password, new_password):
    """修改密碼"""
    db = get_db()
    user = db.users.find_one({"username": username})
    if not user:
        return False, "使用者不存在"
        
    if not bcrypt.checkpw(old_password.encode('utf-8'), user['password']):
        return False, "舊密碼錯誤"
        
    hashed = bcrypt.hashpw(new_password.encode('utf-8'), bcrypt.gensalt())
    db.users.update_one(
        {"username": username},
        {"$set": {"password": hashed}}
    )
    return True, "密碼修改成功"

def get_users():
    """獲取所有使用者"""
    db = get_db()
    return list(db.users.find({}, {"password": 0}))

def delete_user(username):
    """刪除使用者"""
    db = get_db()
    if username == "admin":
        return False, "無法刪除管理員帳號"
    result = db.users.delete_one({"username": username})
    return result.deleted_count > 0, "使用者刪除成功" if result.deleted_count > 0 else "使用者不存在"

def get_org_info():
    """獲取組織資訊"""
    db = get_db()
    org_info = db.org_info.find_one({"_id": 1})
    return org_info or {"org_name": "課堂簽到系統", "manager": "", "contact": ""}

def update_org_info(org_name, manager, contact):
    """更新組織資訊"""
    db = get_db()
    db.org_info.update_one(
        {"_id": 1},
        {"$set": {
            "org_name": org_name,
            "manager": manager,
            "contact": contact
        }},
        upsert=True
    )

def add_class(name):
    """新增課程"""
    db = get_db()
    result = db.classes.insert_one({"name": name})
    return result.inserted_id

def get_classes():
    """獲取所有課程"""
    db = get_db()
    return list(db.classes.find())

def add_session(class_id, week, date, start_time, end_time):
    """新增課程週次"""
    db = get_db()
    result = db.sessions.insert_one({
        "class_id": class_id,
        "week": week,
        "date": date,
        "start_time": start_time,
        "end_time": end_time
    })
    return result.inserted_id

def get_sessions(class_id):
    """獲取課程的所有週次"""
    db = get_db()
    return list(db.sessions.find({"class_id": class_id}).sort("week", 1))

def add_attendee(class_id, name, department, hash_value):
    """新增學員"""
    db = get_db()
    result = db.attendees.insert_one({
        "class_id": class_id,
        "name": name,
        "department": department,
        "hash": hash_value
    })
    return result.inserted_id

def get_attendees(class_id):
    """獲取課程的所有學員"""
    db = get_db()
    return list(db.attendees.find({"class_id": class_id}).sort("name", 1))

def delete_attendee(attendee_id):
    """刪除學員"""
    db = get_db()
    db.attendees.delete_one({"_id": attendee_id})

def check_in(session_id, attendee_id, check_in_time):
    """簽到"""
    db = get_db()
    
    # 檢查學員是否屬於該課程
    session = db.sessions.find_one({"_id": session_id})
    if not session:
        raise ValueError("找不到指定的課程週次")
        
    attendee = db.attendees.find_one({"_id": attendee_id})
    if not attendee:
        raise ValueError("找不到指定的學員")
        
    # 檢查是否已經簽到
    existing_checkin = db.checkins.find_one({
        "session_id": session_id,
        "attendee_id": attendee_id
    })
    
    if existing_checkin and existing_checkin.get("check_in_time"):
        raise ValueError("該學員已經簽到")
    
    # 執行簽到
    db.checkins.update_one(
        {"session_id": session_id, "attendee_id": attendee_id},
        {
            "$set": {
                "check_in_time": check_in_time,
                "class_id": session["class_id"],
                "attendee_name": attendee["name"],
                "department": attendee["department"]
            }
        },
        upsert=True
    )

def check_out(session_id, attendee_id, check_out_time):
    """簽退"""
    db = get_db()
    
    # 檢查是否已經簽到
    checkin = db.checkins.find_one({
        "session_id": session_id,
        "attendee_id": attendee_id
    })
    
    if not checkin or not checkin.get("check_in_time"):
        raise ValueError("該學員尚未簽到")
        
    if checkin.get("check_out_time"):
        raise ValueError("該學員已經簽退")
    
    # 執行簽退
    db.checkins.update_one(
        {"session_id": session_id, "attendee_id": attendee_id},
        {"$set": {"check_out_time": check_out_time}}
    )

def get_checkins(session_id):
    """獲取簽到記錄"""
    db = get_db()
    return list(db.checkins.find({"session_id": session_id}).sort("check_in_time", 1))

def get_attendee_checkin(session_id, attendee_id):
    """獲取特定學員的簽到記錄"""
    db = get_db()
    return db.checkins.find_one({
        "session_id": session_id,
        "attendee_id": attendee_id
    })

def get_attendee_by_hash(hash_value):
    """根據 hash 值獲取學員資訊"""
    db = get_db()
    return db.attendees.find_one({"hash": hash_value})

def update_attendee(attendee_id, name, department):
    """更新學員資料"""
    db = get_db()
    db.attendees.update_one(
        {"_id": attendee_id},
        {"$set": {
            "name": name,
            "department": department,
            "hash": hash_name(name)
        }}
    )

def hash_name(name):
    """根據學員姓名生成唯一的 hash 值"""
    return hashlib.sha256(f"{name}{QR_SEED}".encode()).hexdigest()

def update_default_passwords(admin_password, user_password):
    """更新預設密碼"""
    settings = load_settings()
    settings["admin_password"] = admin_password
    settings["user_password"] = user_password
    save_settings(settings)
    return True, "預設密碼更新成功"

def get_default_passwords():
    """獲取預設密碼"""
    settings = load_settings()
    return {
        "admin_password": settings.get("admin_password", DEFAULT_ADMIN_PASSWORD),
        "user_password": settings.get("user_password", DEFAULT_USER_PASSWORD)
    } 