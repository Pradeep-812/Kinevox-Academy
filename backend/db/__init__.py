"""
db/database.py
==============
SQLite schema + seed data matching the frontend's in-memory state.
"""

import sqlite3
import os
import json
import hashlib
from datetime import datetime


DB_PATH = os.path.join(os.path.dirname(__file__), "..", "kinevox.db")


def get_db():
    """Return a database connection (call per request or use app_context)."""
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    conn.execute("PRAGMA journal_mode = WAL")
    return conn


def _hash_password(pw: str) -> str:
    return hashlib.sha256(pw.encode()).hexdigest()


# ── Schema ──────────────────────────────────────────────────────────────────

SCHEMA = """
-- Users (admin accounts)
CREATE TABLE IF NOT EXISTS users (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    username    TEXT    UNIQUE NOT NULL,
    password    TEXT    NOT NULL,
    role        TEXT    NOT NULL DEFAULT 'admin',
    created_at  TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- Employees
CREATE TABLE IF NOT EXISTS employees (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    role        TEXT,
    dept        TEXT,
    email       TEXT,
    phone       TEXT,
    doj         TEXT,
    salary      INTEGER DEFAULT 0,
    bonus       INTEGER DEFAULT 0,
    credits     INTEGER DEFAULT 0,
    schedule    TEXT,
    status      TEXT    DEFAULT 'active',
    photo       TEXT,
    sign        TEXT,
    id_doc      TEXT,
    resume      TEXT,
    courses     TEXT    DEFAULT '[]',
    created_at  TEXT    NOT NULL DEFAULT (datetime('now')),
    updated_at  TEXT    NOT NULL DEFAULT (datetime('now'))
);

-- Employee Attendance (monthly records)
CREATE TABLE IF NOT EXISTS attendance (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id      TEXT NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    month       TEXT NOT NULL,   -- e.g. "2024-05"
    present     INTEGER DEFAULT 0,
    absent      INTEGER DEFAULT 0,
    late        INTEGER DEFAULT 0,
    UNIQUE(emp_id, month)
);

-- Attendance daily log
CREATE TABLE IF NOT EXISTS attendance_log (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id      TEXT NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    date        TEXT NOT NULL,
    status      TEXT NOT NULL DEFAULT 'present',  -- present | absent | late | off
    note        TEXT,
    UNIQUE(emp_id, date)
);

-- Courses
CREATE TABLE IF NOT EXISTS courses (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    category    TEXT,
    duration    TEXT,
    fee         INTEGER DEFAULT 0,
    seats       INTEGER DEFAULT 0,
    enrolled    INTEGER DEFAULT 0,
    assigned_to TEXT REFERENCES employees(id) ON DELETE SET NULL,
    description TEXT,
    status      TEXT DEFAULT 'active',
    created_at  TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at  TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Students
CREATE TABLE IF NOT EXISTS students (
    id          TEXT PRIMARY KEY,
    name        TEXT NOT NULL,
    email       TEXT,
    phone       TEXT,
    dob         TEXT,
    institution TEXT,
    course      TEXT,
    assigned_to TEXT REFERENCES employees(id) ON DELETE SET NULL,
    trainer     TEXT,
    enroll_date TEXT,
    progress    INTEGER DEFAULT 0,
    percentage  INTEGER DEFAULT 0,
    status      TEXT DEFAULT 'active',
    completed   INTEGER DEFAULT 0,
    photo       TEXT,
    created_at  TEXT NOT NULL DEFAULT (datetime('now')),
    updated_at  TEXT NOT NULL DEFAULT (datetime('now'))
);

-- Certificates (issued records)
CREATE TABLE IF NOT EXISTS certificates (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    student_id  TEXT NOT NULL REFERENCES students(id) ON DELETE CASCADE,
    cert_no     TEXT UNIQUE NOT NULL,
    issued_on   TEXT NOT NULL DEFAULT (datetime('now')),
    course      TEXT NOT NULL
);

-- Salary slip records (history)
CREATE TABLE IF NOT EXISTS salary_slips (
    id          INTEGER PRIMARY KEY AUTOINCREMENT,
    emp_id      TEXT NOT NULL REFERENCES employees(id) ON DELETE CASCADE,
    month       TEXT NOT NULL,
    year        INTEGER NOT NULL,
    basic       INTEGER,
    bonus       INTEGER,
    pf          INTEGER,
    tax         INTEGER,
    gross       INTEGER,
    net         INTEGER,
    generated_on TEXT NOT NULL DEFAULT (datetime('now'))
);
"""

# ── Seed data ────────────────────────────────────────────────────────────────

SEED_USERS = [
    ("admin", "Kinevox123", "admin"),
]

SEED_EMPLOYEES = [
    {
        "id": "EMP001",
        "name": "Arjun Sharma",
        "role": "Senior Developer & Trainer",
        "dept": "Programming",
        "email": "arjun@kinevoxacademy.in",
        "phone": "9876543210",
        "doj": "2023-01-15",
        "salary": 45000,
        "bonus": 5000,
        "credits": 120,
        "schedule": "Mon-Fri 9AM-5PM",
        "status": "active",
        "courses": json.dumps(["Web Development", "React Basics"]),
    },
    {
        "id": "EMP002",
        "name": "Priya Nair",
        "role": "Design & Media Trainer",
        "dept": "Design",
        "email": "priya@kinevoxacademy.in",
        "phone": "9123456789",
        "doj": "2023-03-10",
        "salary": 38000,
        "bonus": 3000,
        "credits": 95,
        "schedule": "Mon-Sat 10AM-6PM",
        "status": "active",
        "courses": json.dumps(["Graphic Design", "Video Editing"]),
    },
    {
        "id": "EMP003",
        "name": "Ravi Menon",
        "role": "Game Dev Instructor",
        "dept": "Programming",
        "email": "ravi@kinevoxacademy.in",
        "phone": "9812345678",
        "doj": "2024-01-01",
        "salary": 42000,
        "bonus": 2000,
        "credits": 60,
        "schedule": "Tue-Sat 11AM-7PM",
        "status": "active",
        "courses": json.dumps(["Game Development"]),
    },
]

SEED_ATTENDANCE = [
    ("EMP001", "2024-05", 22, 2, 1),
    ("EMP002", "2024-05", 20, 3, 2),
    ("EMP003", "2024-05", 18, 1, 3),
]

SEED_COURSES = [
    {
        "id": "CRS001",
        "name": "Web Development",
        "category": "Programming",
        "duration": "3 months",
        "fee": 8000,
        "seats": 15,
        "enrolled": 8,
        "assigned_to": "EMP001",
        "description": "HTML, CSS, JavaScript, React — full-stack fundamentals.",
        "status": "active",
    },
    {
        "id": "CRS002",
        "name": "Graphic Design",
        "category": "Design",
        "duration": "2 months",
        "fee": 6000,
        "seats": 12,
        "enrolled": 6,
        "assigned_to": "EMP002",
        "description": "Photoshop, Illustrator, Canva — visual design mastery.",
        "status": "active",
    },
    {
        "id": "CRS003",
        "name": "Video Editing",
        "category": "Design",
        "duration": "2 months",
        "fee": 7000,
        "seats": 10,
        "enrolled": 4,
        "assigned_to": "EMP002",
        "description": "Premiere Pro, After Effects, DaVinci Resolve.",
        "status": "active",
    },
    {
        "id": "CRS004",
        "name": "React Basics",
        "category": "Programming",
        "duration": "1.5 months",
        "fee": 5500,
        "seats": 12,
        "enrolled": 5,
        "assigned_to": "EMP001",
        "description": "Modern React with hooks, context and state management.",
        "status": "active",
    },
    {
        "id": "CRS005",
        "name": "Game Development",
        "category": "Game Dev",
        "duration": "4 months",
        "fee": 12000,
        "seats": 10,
        "enrolled": 3,
        "assigned_to": "EMP003",
        "description": "Unity 2D/3D game development with C#.",
        "status": "active",
    },
]

SEED_STUDENTS = [
    {
        "id": "STU001",
        "name": "Rahul Verma",
        "email": "rahul@gmail.com",
        "phone": "8765432109",
        "dob": "2005-06-12",
        "institution": "DPS School",
        "course": "Web Development",
        "assigned_to": "EMP001",
        "trainer": "Arjun Sharma",
        "enroll_date": "2024-01-10",
        "progress": 75,
        "percentage": 75,
        "status": "active",
        "completed": 0,
    },
    {
        "id": "STU002",
        "name": "Sneha Patel",
        "email": "sneha@gmail.com",
        "phone": "8654321098",
        "dob": "2004-09-22",
        "institution": "BMS College",
        "course": "Graphic Design",
        "assigned_to": "EMP002",
        "trainer": "Priya Nair",
        "enroll_date": "2024-02-05",
        "progress": 90,
        "percentage": 90,
        "status": "active",
        "completed": 0,
    },
    {
        "id": "STU003",
        "name": "Kiran Kumar",
        "email": "kiran@gmail.com",
        "phone": "9087654321",
        "dob": "2003-12-01",
        "institution": "RV College",
        "course": "Video Editing",
        "assigned_to": "EMP002",
        "trainer": "Priya Nair",
        "enroll_date": "2023-10-01",
        "progress": 100,
        "percentage": 100,
        "status": "completed",
        "completed": 1,
    },
    {
        "id": "STU004",
        "name": "Divya Reddy",
        "email": "divya@gmail.com",
        "phone": "9765432108",
        "dob": "2004-03-15",
        "institution": "Jyoti Nivas",
        "course": "React Basics",
        "assigned_to": "EMP001",
        "trainer": "Arjun Sharma",
        "enroll_date": "2024-03-01",
        "progress": 45,
        "percentage": 45,
        "status": "active",
        "completed": 0,
    },
]


# ── Init ──────────────────────────────────────────────────────────────────────

def init_db(app):
    """Create schema and seed initial data (idempotent)."""
    with app.app_context():
        conn = get_db()
        try:
            # Create tables
            conn.executescript(SCHEMA)

            # Users
            for username, password, role in SEED_USERS:
                conn.execute(
                    "INSERT OR IGNORE INTO users (username, password, role) VALUES (?,?,?)",
                    (username, _hash_password(password), role),
                )

            # Employees
            for e in SEED_EMPLOYEES:
                conn.execute(
                    """INSERT OR IGNORE INTO employees
                       (id,name,role,dept,email,phone,doj,salary,bonus,credits,
                        schedule,status,courses)
                       VALUES (:id,:name,:role,:dept,:email,:phone,:doj,:salary,:bonus,
                               :credits,:schedule,:status,:courses)""",
                    e,
                )

            # Attendance summary
            for emp_id, month, present, absent, late in SEED_ATTENDANCE:
                conn.execute(
                    """INSERT OR IGNORE INTO attendance (emp_id,month,present,absent,late)
                       VALUES (?,?,?,?,?)""",
                    (emp_id, month, present, absent, late),
                )

            # Courses
            for c in SEED_COURSES:
                conn.execute(
                    """INSERT OR IGNORE INTO courses
                       (id,name,category,duration,fee,seats,enrolled,
                        assigned_to,description,status)
                       VALUES (:id,:name,:category,:duration,:fee,:seats,:enrolled,
                               :assigned_to,:description,:status)""",
                    c,
                )

            # Students
            for s in SEED_STUDENTS:
                conn.execute(
                    """INSERT OR IGNORE INTO students
                       (id,name,email,phone,dob,institution,course,assigned_to,trainer,
                        enroll_date,progress,percentage,status,completed)
                       VALUES (:id,:name,:email,:phone,:dob,:institution,:course,
                               :assigned_to,:trainer,:enroll_date,:progress,:percentage,
                               :status,:completed)""",
                    s,
                )

            conn.commit()
        finally:
            conn.close()
