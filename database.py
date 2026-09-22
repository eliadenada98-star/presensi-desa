import sqlite3
from pathlib import Path
from werkzeug.security import generate_password_hash


# =========================================================
# KONFIGURASI DATABASE
# =========================================================

BASE_DIR = Path(__file__).resolve().parent

DATABASE = BASE_DIR / "presensi.db"


# =========================================================
# KONEKSI DATABASE
# =========================================================

def get_connection():

    conn = sqlite3.connect(
        DATABASE
    )

    conn.row_factory = sqlite3.Row

    conn.execute(
        "PRAGMA foreign_keys = ON"
    )

    return conn


# =========================================================
# EXECUTE QUERY
# =========================================================

def execute(
    query,
    params=()
):

    conn = get_connection()

    try:

        cursor = conn.execute(
            query,
            params
        )

        conn.commit()

        return cursor.lastrowid

    finally:

        conn.close()


# =========================================================
# AMBIL SATU DATA
# =========================================================

def fetch_one(
    query,
    params=()
):

    conn = get_connection()

    try:

        return conn.execute(
            query,
            params
        ).fetchone()

    finally:

        conn.close()


# =========================================================
# AMBIL SEMUA DATA
# =========================================================

def fetch_all(
    query,
    params=()
):

    conn = get_connection()

    try:

        return conn.execute(
            query,
            params
        ).fetchall()

    finally:

        conn.close()


# =========================================================
# TAMBAH KOLOM JIKA BELUM ADA
# =========================================================

def add_column_if_not_exists(
    conn,
    table_name,
    column_name,
    column_definition
):

    columns = conn.execute(
        f"PRAGMA table_info({table_name})"
    ).fetchall()

    existing_columns = [
        column["name"]
        for column in columns
    ]

    if column_name not in existing_columns:

        conn.execute(
            f"""
            ALTER TABLE {table_name}
            ADD COLUMN {column_name}
            {column_definition}
            """
        )


# =========================================================
# INISIALISASI DATABASE
# =========================================================

def init_database():

    conn = get_connection()

    cursor = conn.cursor()

    # =====================================================
    # TABEL PERANGKAT
    # =====================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS perangkat (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            nama TEXT NOT NULL,

            nik TEXT,

            jabatan TEXT NOT NULL,

            foto TEXT,

            status TEXT NOT NULL
                DEFAULT 'Aktif',

            created_at TIMESTAMP
                DEFAULT CURRENT_TIMESTAMP

        )
    """)


    # =====================================================
    # TABEL USERS
    # =====================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS users (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            username TEXT NOT NULL UNIQUE,

            password TEXT NOT NULL,

            nama TEXT NOT NULL,

            role TEXT NOT NULL
                DEFAULT 'perangkat',

            perangkat_id INTEGER,

            created_at TIMESTAMP
                DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (perangkat_id)
                REFERENCES perangkat(id)

                ON DELETE SET NULL

        )
    """)


    # =====================================================
    # TABEL PRESENSI
    # =====================================================

    cursor.execute("""
        CREATE TABLE IF NOT EXISTS presensi (

            id INTEGER PRIMARY KEY AUTOINCREMENT,

            perangkat_id INTEGER NOT NULL,

            tanggal TEXT NOT NULL,

            jam_masuk TEXT,

            foto_masuk TEXT,

            jam_pulang TEXT,

            foto_pulang TEXT,

            latitude REAL,

            longitude REAL,

            status TEXT NOT NULL
                DEFAULT 'Hadir',

            created_at TIMESTAMP
                DEFAULT CURRENT_TIMESTAMP,

            FOREIGN KEY (perangkat_id)
                REFERENCES perangkat(id)

                ON DELETE CASCADE

        )
    """)


    # =====================================================
    # MIGRASI GPS
    # =====================================================

    add_column_if_not_exists(
        conn,
        "presensi",
        "latitude_masuk",
        "REAL"
    )


    add_column_if_not_exists(
        conn,
        "presensi",
        "longitude_masuk",
        "REAL"
    )


    add_column_if_not_exists(
        conn,
        "presensi",
        "latitude_pulang",
        "REAL"
    )


    add_column_if_not_exists(
        conn,
        "presensi",
        "longitude_pulang",
        "REAL"
    )


    # =====================================================
    # PASTIKAN KOLOM FOTO PERANGKAT ADA
    # =====================================================

    add_column_if_not_exists(
        conn,
        "perangkat",
        "foto",
        "TEXT"
    )


    # =====================================================
    # AKUN ADMIN DEFAULT
    # =====================================================

    admin = cursor.execute("""
        SELECT id

        FROM users

        WHERE username = ?

    """, (
        "admin",
    )).fetchone()


    if admin is None:

        password_hash = (
            generate_password_hash(
                "admin123"
            )
        )

        cursor.execute("""
            INSERT INTO users
            (
                username,
                password,
                nama,
                role,
                perangkat_id
            )

            VALUES
            (
                ?,
                ?,
                ?,
                ?,
                ?
            )

        """, (
            "admin",
            password_hash,
            "Administrator",
            "admin",
            None
        ))


    # =====================================================
    # SIMPAN PERUBAHAN
    # =====================================================

    conn.commit()

    conn.close()


# =========================================================
# JALANKAN DATABASE.PY LANGSUNG
# =========================================================

if __name__ == "__main__":

    init_database()

    print()
    print(
        "========================================"
    )

    print(
        " DATABASE BERHASIL DIBUAT / DIPERBARUI"
    )

    print(
        "========================================"
    )

    print()

    print(
        "Akun Admin"
    )

    print(
        "Username : admin"
    )

    print(
        "Password : admin123"
    )

    print()