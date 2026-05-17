from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse, FileResponse
from fastapi.staticfiles import StaticFiles
from pydantic import BaseModel
import requests
import psycopg2
import psycopg2.extras
import os
import json
import re
import io
from typing import Optional
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment

app = FastAPI(title="Refinery Contract Chatbot API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_methods=["*"],
    allow_headers=["*"],
)

# Static files & root route
app.mount("/static", StaticFiles(directory="static"), name="static")

@app.get("/")
def root():
    return FileResponse("static/index.html")

# ── Config ───────────────────────────────────────────────────────────────────
DATABASE_URL = os.getenv("DATABASE_URL", "postgresql://postgres:password@host:5432/railway")
DINOIKI_API_KEY = os.getenv("DINOIKI_API_KEY", "")
DINOIKI_URL = "https://ai.dinoiki.com/v1/chat/completions"
AI_MODEL = "gpt-4o"

# ── Stopwords bahasa Indonesia ────────────────────────────────────────────────
STOPWORDS = {
    'apa', 'siapa', 'berapa', 'yang', 'dan', 'atau', 'dari', 'untuk',
    'dengan', 'adalah', 'ini', 'itu', 'ada', 'tidak', 'bisa', 'mau',
    'saya', 'kamu', 'dia', 'kami', 'kita', 'mereka', 'semua', 'sudah',
    'belum', 'sedang', 'akan', 'telah', 'pada', 'di', 'ke', 'oleh',
    'juga', 'hanya', 'lebih', 'paling', 'sangat', 'banyak', 'sedikit',
    'tampilkan', 'tunjukkan', 'cari', 'lihat', 'data', 'info', 'informasi',
    'list', 'daftar', 'total', 'jumlah', 'nilai', 'status', 'semua',
    'kontrak', 'vendor', 'tagihan', 'dokumen', 'progress', 'bulan', 'tahun'
}

# ── Helper: call dinoiki AI ──────────────────────────────────────────────────
def call_ai(messages: list, max_tokens: int = 1500) -> str:
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {DINOIKI_API_KEY}"
    }
    payload = {
        "model": AI_MODEL,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": 0.3,
    }
    resp = requests.post(DINOIKI_URL, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"].strip()

# ── Dynamic Context Injection ─────────────────────────────────────────────────
def smart_entity_search(user_message: str) -> str:
    """
    Cari entitas yang disebut user secara dinamis di database.
    Hasilnya diinjeksikan ke system prompt agar AI tahu konteksnya.
    """
    try:
        conn = psycopg2.connect(DATABASE_URL)
        cur = conn.cursor()
        context = []
        found_ids = set()  # hindari duplikasi

        # Ekstrak kata-kata penting dari pesan user
        words = [w for w in re.findall(r'\b\w{3,}\b', user_message)
                 if w.lower() not in STOPWORDS]

        # Juga cari frasa PT/CV/UD secara khusus
        company_patterns = re.findall(r'\b(?:PT|CV|UD|TB|PD)\s+[\w\s]+', user_message, re.IGNORECASE)

        search_terms = list(set(words + company_patterns))

        for term in search_terms:
            term = term.strip()
            if len(term) < 3:
                continue

            # ── Cari di tabel vendor ──────────────────────────────
            cur.execute("""
                SELECT id_vendor, nama_vendor, status_vendor, score
                FROM vendor
                WHERE nama_vendor ILIKE %s
                LIMIT 3
            """, (f'%{term}%',))
            vendors = cur.fetchall()
            for v in vendors:
                key = f"vendor_{v[0]}"
                if key not in found_ids:
                    found_ids.add(key)
                    context.append(
                        f"[VENDOR] '{v[1]}' → id_vendor={v[0]}, "
                        f"status={v[2]}, score={v[3]}"
                    )

            # ── Cari di tabel kontrak ─────────────────────────────
            cur.execute("""
                SELECT k.id_kontrak, k.judul_kontrak, k.no_dokumen_kontrak,
                       k.direksi_pekerjaan, k.status_kontrak, k.tipe_kontrak,
                       v.nama_vendor
                FROM kontrak k
                LEFT JOIN vendor v ON k.id_vendor = v.id_vendor
                WHERE k.judul_kontrak ILIKE %s
                   OR k.no_dokumen_kontrak ILIKE %s
                   OR k.no_po_pr ILIKE %s
                   OR k.direksi_pekerjaan ILIKE %s
                LIMIT 3
            """, (f'%{term}%', f'%{term}%', f'%{term}%', f'%{term}%'))
            kontraks = cur.fetchall()
            for k in kontraks:
                key = f"kontrak_{k[0]}"
                if key not in found_ids:
                    found_ids.add(key)
                    context.append(
                        f"[KONTRAK] '{k[1]}' → id_kontrak={k[0]}, "
                        f"doc={k[2]}, direksi={k[3]}, "
                        f"status={k[4]}, tipe={k[5]}, vendor='{k[6]}'"
                    )

            # ── Cari di tabel tagihan ─────────────────────────────
            cur.execute("""
                SELECT t.id_tagihan, t.nomor_tagihan, t.status_tagihan,
                       t.nilai_tagihan, k.judul_kontrak
                FROM tagihan t
                LEFT JOIN kontrak k ON t.id_kontrak = k.id_kontrak
                WHERE t.nomor_tagihan ILIKE %s
                LIMIT 3
            """, (f'%{term}%',))
            taghans = cur.fetchall()
            for t in taghans:
                key = f"tagihan_{t[0]}"
                if key not in found_ids:
                    found_ids.add(key)
                    context.append(
                        f"[TAGIHAN] '{t[1]}' → id_tagihan={t[0]}, "
                        f"status={t[2]}, nilai={t[3]}, kontrak='{t[4]}'"
                    )

            # ── Cari di tabel padi ────────────────────────────────
            cur.execute("""
                SELECT p.id_padi, p.no_pembelian, p.judul_pembelian,
                       p.nilai, v.nama_vendor
                FROM padi p
                LEFT JOIN vendor v ON p.id_vendor = v.id_vendor
                WHERE p.no_pembelian ILIKE %s
                   OR p.judul_pembelian ILIKE %s
                LIMIT 3
            """, (f'%{term}%', f'%{term}%'))
            padis = cur.fetchall()
            for p in padis:
                key = f"padi_{p[0]}"
                if key not in found_ids:
                    found_ids.add(key)
                    context.append(
                        f"[PADI] '{p[2]}' → id_padi={p[0]}, "
                        f"no_pembelian={p[1]}, nilai={p[3]}, vendor='{p[4]}'"
                    )

        conn.close()

        if context:
            result = "\n\nKONTEKS ENTITAS YANG DITEMUKAN DI DATABASE:\n"
            result += "(Gunakan informasi ini untuk memahami maksud user tanpa perlu klarifikasi)\n"
            result += "\n".join(context)
            return result

        return ""

    except Exception as e:
        # Jangan crash jika search gagal, cukup return kosong
        return ""

# ── DB Schema context untuk AI ───────────────────────────────────────────────
SCHEMA_CONTEXT = """
Database PostgreSQL untuk sistem manajemen kontrak kilang minyak. Berikut skema tabel:

TABEL: profiles
Kolom: id, email, full_name, role (admin/pic/user), password_hash, created_at, updated_at, is_active, id_vendor

TABEL: vendor
Kolom: id_vendor, nama_vendor, npwp, alamat, pic_nama, pic_kontak, status_vendor (Active/Inactive/Blacklist), score, created_at, updated_at

TABEL: kontrak
Kolom: id_kontrak, id_vendor, judul_kontrak, no_dokumen_kontrak, no_po_pr, direksi_pekerjaan,
  tipe_kontrak (Lumpsum/Unit Price/TSA/LTSA/TSA-LTSA), status_kontrak (Pre-KOM/Active/Aktif/Completed/Selesai/Terminated),
  tanggal_spb_diterima, tanggal_terima_dokumen, tanggal_maksimal_kom, tanggal_mulai, tanggal_selesai,
  sla_kom_hari, estimasi_tanggal_kom, tanggal_kom, kom_terlambat, nilai_awal, durasi_kontrak_hari,
  progress_plan, progress_actual, aktivitas_saat_ini, kendala, disiplin, tkdn_percentage, tanggal_lkp,
  has_amendment, no_amandemen, tanggal_amandemen, jenis_amandemen, nilai_kontrak_baru, durasi_amandemen,
  tanggal_mulai_baru, tanggal_selesai_baru, alasan_perubahan, contract_documents, amendment_documents,
  s_curve_data, created_at, updated_at

TABEL: amandemen_kontrak
Kolom: id_amandemen, id_kontrak, nomor_urut, no_amandemen, tanggal_amandemen, jenis_amandemen,
  nilai_kontrak_baru, durasi_amandemen, tanggal_mulai_baru, tanggal_selesai_baru, alasan_perubahan,
  amendment_documents, created_at, updated_at

TABEL: tagihan
Kolom: id_tagihan, id_kontrak, nomor_tagihan, tanggal_tagihan, tipe_kontrak, termin, nilai_tagihan,
  status_tagihan, memo_required, tanggal_pengiriman_memo, dokumen_memo, dokumen_tagihan, catatan,
  created_at, updated_at

TABEL: progress_lumpsum
Kolom: id_progress, id_kontrak, milestone, persen, tanggal_update, evidence, created_at

TABEL: progress_unit_price
Kolom: id_progress, id_kontrak, nama_item, satuan, qty_rencana, qty_aktual, harga_satuan, tanggal_update, created_at

TABEL: monitoring_ltsa
Kolom: id_log, id_kontrak, tanggal_kunjungan, jenis_layanan (Preventive/Corrective/Standby),
  durasi_jam, sla_terpenuhi (Yes/No), keterangan, created_at

TABEL: padi
Kolom: id_padi, no_pembelian, tanggal, judul_pembelian, no_po_pr, nilai, id_vendor, link_pembelian,
  bagian, dokumen_pendukung, status_purchase (BAST), tanggal_bast, tanggal_sa_gr, tanggal_invoice,
  tanggal_payment_approval, tanggal_paid, catatan_status, created_at, updated_at

TABEL: dokumen_approval
Kolom: id_dokumen, id_kontrak, tipe_dokumen (Evident/Report/Persetujuan), nama_dokumen,
  deskripsi_dokumen, file_path, file_url, nama_file, tipe_file, ukuran_file,
  status_approval (Pending/Approved/Rejected), catatan_reviewer, uploaded_by, reviewed_by,
  reviewed_at, created_at, updated_at

TABEL: konfigurasi_sistem
Kolom: id_setting, nama_setting, nilai_setting, deskripsi, updated_at

Relasi penting:
- vendor.id_vendor → kontrak.id_vendor (1 vendor banyak kontrak)
- kontrak.id_kontrak → tagihan.id_kontrak
- kontrak.id_kontrak → amandemen_kontrak.id_kontrak
- kontrak.id_kontrak → progress_lumpsum.id_kontrak
- kontrak.id_kontrak → progress_unit_price.id_kontrak
- kontrak.id_kontrak → monitoring_ltsa.id_kontrak
- kontrak.id_kontrak → dokumen_approval.id_kontrak
- vendor.id_vendor → padi.id_vendor

NILAI ENUM & PILIHAN YANG VALID:

1. TIPE KONTRAK (kontrak.tipe_kontrak):
   - 'Lumpsum'     → kontrak dengan nilai tetap, pembayaran per milestone
   - 'Unit Price'  → kontrak berdasarkan satuan pekerjaan/volume
   - 'TSA'         → Technical Service Agreement, jasa teknis rutin
   - 'LTSA'        → Long Term Service Agreement, jasa jangka panjang
   - 'TSA/LTSA'    → gabungan TSA dan LTSA

2. STATUS KONTRAK (kontrak.status_kontrak):
   - 'Pre-KOM'     → kontrak belum mulai, masih tahap persiapan KOM
   - 'Aktif'       → kontrak sedang berjalan
   - 'Selesai'     → kontrak telah selesai dilaksanakan
   - 'Terminated'  → kontrak dihentikan sebelum selesai

3. DISIPLIN (kontrak.disiplin):
   - 'Instrumentasi' → pekerjaan instrumen & kontrol
   - 'Stationary'    → pekerjaan bejana/vessel statis
   - 'Electrical'    → pekerjaan kelistrikan
   - 'Rotating'      → pekerjaan mesin berputar (pompa, kompresor)
   - 'Alat Berat'    → pekerjaan menggunakan alat berat

4. DIREKSI PEKERJAAN (kontrak.direksi_pekerjaan):
   - 'MA5'       → unit/departemen MA5
   - 'MA6'       → unit/departemen MA6
   - 'MA7'       → unit/departemen MA7
   - 'Workshop'  → unit Workshop

5. JENIS AMANDEMEN (amandemen_kontrak.jenis_amandemen):
   - 'Nilai'          → perubahan nilai kontrak saja
   - 'Waktu'          → perubahan durasi/tanggal saja
   - 'Nilai dan Waktu' → perubahan nilai sekaligus durasi

6. TIPE DOKUMEN APPROVAL (dokumen_approval.tipe_dokumen):
   - 'Evident Progress'       → bukti progress pekerjaan
   - 'Report Vendor'          → laporan dari vendor
   - 'Persetujuan Pelaksanaan' → dokumen persetujuan pelaksanaan

7. STATUS APPROVAL (dokumen_approval.status_approval):
   - 'Pending'  → menunggu review
   - 'Approved' → sudah disetujui
   - 'Rejected' → ditolak

8. STATUS VENDOR (vendor.status_vendor):
   - 'Active'    → vendor aktif dan dapat digunakan
   - 'Inactive'  → vendor tidak aktif
   - 'Blacklist' → vendor diblacklist, tidak boleh digunakan

9. JENIS LAYANAN LTSA (monitoring_ltsa.jenis_layanan):
   - 'Preventive' → perawatan rutin/pencegahan
   - 'Corrective' → perbaikan kerusakan
   - 'Standby'    → siaga/standby

10. SLA TERPENUHI (monitoring_ltsa.sla_terpenuhi):
    - 'Yes' → SLA terpenuhi
    - 'No'  → SLA tidak terpenuhi

11. STATUS TAGIHAN (tagihan.status_tagihan) — mengikuti tahapan progress:
    Tahap 1: Punchlist
    Tahap 2: BAST/BAPP
    Tahap 3: Pengajuan
    Tahap 4: BAST I Vendor
    Tahap 5: SA (Service Acceptance)
    Tahap 6: PA (Payment Approval)
    Tahap 7: Verification
    Tahap 8: Payment/Selesai

12. STATUS PURCHASE PADI (padi.status_purchase):
    - 'BAST' → Berita Acara Serah Terima sudah dilakukan
"""
- vendor.id_vendor → kontrak.id_vendor (1 vendor banyak kontrak)
- kontrak.id_kontrak → tagihan.id_kontrak
- kontrak.id_kontrak → amandemen_kontrak.id_kontrak
- kontrak.id_kontrak → progress_lumpsum.id_kontrak
- kontrak.id_kontrak → progress_unit_price.id_kontrak
- kontrak.id_kontrak → monitoring_ltsa.id_kontrak
- kontrak.id_kontrak → dokumen_approval.id_kontrak
- vendor.id_vendor → padi.id_vendor
"""

BASE_SYSTEM_PROMPT = f"""Kamu adalah asisten cerdas untuk sistem manajemen kontrak kilang minyak.
Kamu dapat menjawab pertanyaan bisnis dalam bahasa Indonesia secara natural dan mengkonversinya ke query SQL PostgreSQL.

{SCHEMA_CONTEXT}

ATURAN KETAT:
1. HANYA boleh generate query SELECT, TIDAK boleh UPDATE, DELETE, INSERT, DROP, ALTER, TRUNCATE, dll
2. TIDAK boleh query SELECT * (tanpa kolom spesifik) - selalu tentukan kolom yang relevan
3. Selalu gunakan LIMIT maksimal 1000 baris
4. Gunakan JOIN yang tepat antar tabel
5. Format angka nilai kontrak dalam format Indonesia (Rp)

ATURAN INTERPRETASI ENTITAS:
- Jika ada blok "KONTEKS ENTITAS YANG DITEMUKAN DI DATABASE" → gunakan langsung, JANGAN minta klarifikasi
- Jika user menyebut nama yang diawali PT/CV/UD → cari di vendor.nama_vendor
- Jika user menyebut kode seperti MA5, KOM-001, KTR-xxx → cari di direksi_pekerjaan atau no_dokumen_kontrak
- Jika entitas tidak ditemukan di konteks → baru boleh minta klarifikasi

FORMAT RESPONS JSON:
Kamu HARUS selalu merespons dalam format JSON seperti ini:
{{
  "type": "query" | "clarification" | "narrative" | "error",
  "sql": "query SQL jika type=query",
  "explanation": "penjelasan dalam bahasa Indonesia apa yang akan dilakukan query ini",
  "narrative_hint": "bagaimana cara menarasikan hasilnya nanti",
  "chart_suggestion": null | "bar" | "line" | "pie" | "doughnut",
  "chart_config": null | {{"x_column": "...", "y_column": "...", "label": "..."}},
  "clarification_question": "pertanyaan klarifikasi jika type=clarification",
  "message": "pesan untuk user"
}}

DETEKSI CHART:
- Jika pertanyaan menyebut "grafik", "chart", "trend", "perbandingan", "distribusi", "per bulan/tahun" → suggest chart
- bar chart: perbandingan kategori (status, tipe, vendor)
- line chart: data time-series (per bulan, trend progress)
- pie/doughnut: distribusi persentase (status kontrak, tipe)

DETEKSI TABEL:
- Jika hasil query berpotensi >5 baris dan multi-kolom → akan ditampilkan tabel
- Jika hanya 1-2 nilai → tampilkan sebagai narasi saja
"""

# ── Models ────────────────────────────────────────────────────────────────────
class ChatRequest(BaseModel):
    message: str
    history: list = []

class DownloadRequest(BaseModel):
    sql: str
    filename: str = "data_export"

# ── DB Connection ─────────────────────────────────────────────────────────────
def get_db_connection():
    try:
        conn = psycopg2.connect(DATABASE_URL)
        return conn
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Database connection failed: {str(e)}")

# ── SQL Validator ─────────────────────────────────────────────────────────────
def validate_sql(sql: str) -> tuple[bool, str]:
    sql_upper = sql.upper().strip()

    dangerous = ["UPDATE", "DELETE", "INSERT", "DROP", "ALTER", "TRUNCATE", "CREATE", "GRANT", "REVOKE"]
    for op in dangerous:
        pattern = r'\b' + op + r'\b'
        if re.search(pattern, sql_upper):
            return False, f"Operasi {op} tidak diizinkan. Hanya query SELECT yang diperbolehkan."

    if re.search(r'SELECT\s+\*', sql_upper):
        return False, "Query SELECT * tidak diizinkan. Harap tentukan kolom yang spesifik."

    if not sql_upper.startswith("SELECT") and "SELECT" not in sql_upper[:50]:
        return False, "Hanya query SELECT yang diizinkan."

    if "LIMIT" not in sql_upper:
        sql = sql.rstrip(";") + " LIMIT 1000"

    return True, sql

# ── Execute Query ─────────────────────────────────────────────────────────────
def execute_query(sql: str) -> tuple[list, list]:
    valid, result = validate_sql(sql)
    if not valid:
        raise HTTPException(status_code=400, detail=result)

    sql = result

    conn = get_db_connection()
    try:
        with conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor) as cur:
            cur.execute(sql)
            rows = cur.fetchall()
            columns = [desc[0] for desc in cur.description] if cur.description else []
            data = [dict(row) for row in rows]
            return data, columns
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Query error: {str(e)}")
    finally:
        conn.close()

# ── Generate Narrative ────────────────────────────────────────────────────────
def generate_narrative(data: list, columns: list, original_question: str, narrative_hint: str) -> str:
    if not data:
        return "Tidak ditemukan data yang sesuai dengan pertanyaan Anda."

    if len(data) <= 5 and len(columns) <= 5:
        data_str = json.dumps(data, default=str, ensure_ascii=False)
        try:
            return call_ai([
                {
                    "role": "system",
                    "content": "Kamu adalah asisten laporan bisnis. Jawab hanya dalam bahasa Indonesia yang profesional dan natural."
                },
                {
                    "role": "user",
                    "content": (
                        f'Pertanyaan user: "{original_question}"\n'
                        f"Hint narasi: {narrative_hint}\n"
                        f"Data hasil query: {data_str}\n\n"
                        "Buatkan narasi singkat dalam bahasa Indonesia yang menjawab pertanyaan tersebut secara profesional. "
                        "Jika ada nilai uang, format sebagai Rupiah (contoh: Rp 1.250.000.000). Maksimal 3 kalimat."
                    )
                }
            ], max_tokens=400)
        except Exception:
            return ""

    return ""

# ── Chat Endpoint ─────────────────────────────────────────────────────────────
@app.post("/api/chat")
async def chat(req: ChatRequest):
    # 1. Cari entitas yang disebut user di database secara dinamis
    dynamic_context = smart_entity_search(req.message)

    # 2. Gabungkan base prompt + konteks dinamis
    system_prompt = BASE_SYSTEM_PROMPT + dynamic_context

    # 3. Build messages
    messages = [{"role": "system", "content": system_prompt}]
    for h in req.history[-10:]:
        messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": req.message})

    try:
        raw = call_ai(messages, max_tokens=1500)

        json_match = re.search(r'\{[\s\S]*\}', raw)
        if not json_match:
            return {
                "type": "narrative",
                "message": raw,
                "data": None,
                "columns": [],
                "chart": None,
                "narrative": raw,
                "sql": None,
                "row_count": 0
            }

        parsed = json.loads(json_match.group())
        response_type = parsed.get("type", "narrative")

        if response_type == "clarification":
            return {
                "type": "clarification",
                "message": parsed.get("clarification_question", parsed.get("message", "")),
                "data": None,
                "columns": [],
                "chart": None,
                "narrative": None,
                "sql": None,
                "row_count": 0
            }

        if response_type == "error":
            return {
                "type": "error",
                "message": parsed.get("message", "Terjadi kesalahan."),
                "data": None,
                "columns": [],
                "chart": None,
                "narrative": None,
                "sql": None,
                "row_count": 0
            }

        if response_type == "query" and parsed.get("sql"):
            sql = parsed["sql"]

            valid, val_result = validate_sql(sql)
            if not valid:
                return {
                    "type": "error",
                    "message": val_result,
                    "data": None,
                    "columns": [],
                    "chart": None,
                    "narrative": None,
                    "sql": None,
                    "row_count": 0
                }

            data, columns = execute_query(val_result)
            row_count = len(data)

            narrative = ""
            if row_count > 0 and row_count <= 5 and len(columns) <= 5:
                narrative = generate_narrative(
                    data, columns, req.message,
                    parsed.get("narrative_hint", "")
                )

            chart = None
            if parsed.get("chart_suggestion") and row_count > 0:
                chart = {
                    "type": parsed["chart_suggestion"],
                    "config": parsed.get("chart_config", {})
                }

            serializable_data = []
            for row in data:
                clean_row = {}
                for k, v in row.items():
                    if hasattr(v, 'isoformat'):
                        clean_row[k] = v.isoformat()
                    elif v is None:
                        clean_row[k] = None
                    else:
                        clean_row[k] = v
                serializable_data.append(clean_row)

            return {
                "type": "query",
                "message": parsed.get("explanation", ""),
                "data": serializable_data,
                "columns": columns,
                "chart": chart,
                "narrative": narrative,
                "sql": val_result,
                "row_count": row_count
            }

        return {
            "type": "narrative",
            "message": parsed.get("message", raw),
            "data": None,
            "columns": [],
            "chart": None,
            "narrative": parsed.get("message", raw),
            "sql": None,
            "row_count": 0
        }

    except json.JSONDecodeError:
        return {
            "type": "narrative",
            "message": raw,
            "data": None,
            "columns": [],
            "chart": None,
            "narrative": raw,
            "sql": None,
            "row_count": 0
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

# ── Download Excel ────────────────────────────────────────────────────────────
@app.post("/api/download")
async def download_excel(req: DownloadRequest):
    data, columns = execute_query(req.sql)

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Data Export"

    header_fill = PatternFill(start_color="1A2744", end_color="1A2744", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)

    for col_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name.upper().replace("_", " "))
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    alt_fill = PatternFill(start_color="E8EEF8", end_color="E8EEF8", fill_type="solid")
    for row_idx, row in enumerate(data, 2):
        for col_idx, col_name in enumerate(columns, 1):
            val = row.get(col_name)
            cell = ws.cell(row=row_idx, column=col_idx, value=str(val) if val is not None else "")
            if row_idx % 2 == 0:
                cell.fill = alt_fill

    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)

    filename = f"{req.filename}.xlsx"
    return StreamingResponse(
        buf,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.get("/api/health")
def health():
    return {"status": "ok", "message": "Refinery Contract Chatbot API running"}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)