from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import StreamingResponse
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

# ── Config ──────────────────────────────────────────────────────────────────
DATABASE_URL = os.getenv("DATABASE_URL", "postgresql://postgres:password@host:5432/railway")
DINOIKI_API_KEY = os.getenv("DINOIKI_API_KEY", "")
DINOIKI_URL = "https://ai.dinoiki.com/v1/chat/completions"
AI_MODEL = "gpt-4o"

# ── Helper: call dinoiki AI ───────────────────────────────────────────────────
def call_ai(messages: list, max_tokens: int = 1500) -> str:
    """Panggil dinoiki API dan kembalikan teks respons."""
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {DINOIKI_API_KEY}"
    }
    payload = {
        "model": AI_MODEL,
        "messages": messages,
        "max_tokens": max_tokens,
        "temperature": 0.3,   # rendah biar konsisten untuk SQL generation
    }
    resp = requests.post(DINOIKI_URL, headers=headers, json=payload, timeout=60)
    resp.raise_for_status()
    data = resp.json()
    return data["choices"][0]["message"]["content"].strip()

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
"""

SYSTEM_PROMPT = f"""Kamu adalah asisten cerdas untuk sistem manajemen kontrak kilang minyak. 
Kamu dapat menjawab pertanyaan bisnis dalam bahasa Indonesia secara natural dan mengkonversinya ke query SQL PostgreSQL.

{SCHEMA_CONTEXT}

ATURAN KETAT:
1. HANYA boleh generate query SELECT, TIDAK boleh UPDATE, DELETE, INSERT, DROP, ALTER, TRUNCATE, dll
2. TIDAK boleh query SELECT * (tanpa kolom spesifik) - selalu tentukan kolom yang relevan
3. Selalu gunakan LIMIT maksimal 1000 baris
4. Jika pertanyaan tidak jelas atau ambigu, minta klarifikasi
5. Gunakan JOIN yang tepat antar tabel
6. Format angka nilai kontrak dalam format Indonesia (Rp)

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

Jangan pernah generate query tanpa WHERE yang masuk akal (kecuali agregasi summary).
"""

# ── Models ───────────────────────────────────────────────────────────────────
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
    
    # Cek operasi berbahaya
    dangerous = ["UPDATE", "DELETE", "INSERT", "DROP", "ALTER", "TRUNCATE", "CREATE", "GRANT", "REVOKE"]
    for op in dangerous:
        pattern = r'\b' + op + r'\b'
        if re.search(pattern, sql_upper):
            return False, f"Operasi {op} tidak diizinkan. Hanya query SELECT yang diperbolehkan."
    
    # Cek SELECT *
    if re.search(r'SELECT\s+\*', sql_upper):
        return False, "Query SELECT * tidak diizinkan. Harap tentukan kolom yang spesifik."
    
    # Harus ada SELECT
    if not sql_upper.startswith("SELECT") and "SELECT" not in sql_upper[:50]:
        return False, "Hanya query SELECT yang diizinkan."
    
    # Cek LIMIT
    if "LIMIT" not in sql_upper:
        sql = sql.rstrip(";") + " LIMIT 1000"
    
    return True, sql

# ── Execute Query ─────────────────────────────────────────────────────────────
def execute_query(sql: str) -> tuple[list, list]:
    valid, result = validate_sql(sql)
    if not valid:
        raise HTTPException(status_code=400, detail=result)
    
    sql = result  # might have LIMIT appended
    
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
    
    # Untuk data kecil (≤5 baris, ≤5 kolom), buat narasi natural
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
    # Build messages — system prompt sebagai message pertama (role=system)
    messages = [{"role": "system", "content": SYSTEM_PROMPT}]
    for h in req.history[-10:]:  # keep last 10 turns
        messages.append({"role": h["role"], "content": h["content"]})
    messages.append({"role": "user", "content": req.message})
    
    # Get AI response via dinoiki
    try:
        raw = call_ai(messages, max_tokens=1500)
        
        # Parse JSON dari response AI
        # Cari JSON block
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
            
            # Validate & execute
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
            
            # Generate narrative for small datasets
            narrative = ""
            if row_count > 0 and row_count <= 5 and len(columns) <= 5:
                narrative = generate_narrative(
                    data, columns, req.message,
                    parsed.get("narrative_hint", "")
                )
            
            # Prepare chart config
            chart = None
            if parsed.get("chart_suggestion") and row_count > 0:
                chart = {
                    "type": parsed["chart_suggestion"],
                    "config": parsed.get("chart_config", {})
                }
            
            # Convert data for JSON serialization
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
        
        # Fallback narrative
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
    
    # Header style
    header_fill = PatternFill(start_color="1a1a1a", end_color="1a1a1a", fill_type="solid")
    header_font = Font(color="FFFFFF", bold=True, name="Calibri", size=11)
    
    # Write headers
    for col_idx, col_name in enumerate(columns, 1):
        cell = ws.cell(row=1, column=col_idx, value=col_name.upper().replace("_", " "))
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    
    # Write data
    alt_fill = PatternFill(start_color="F5F5F5", end_color="F5F5F5", fill_type="solid")
    for row_idx, row in enumerate(data, 2):
        for col_idx, col_name in enumerate(columns, 1):
            val = row.get(col_name)
            cell = ws.cell(row=row_idx, column=col_idx, value=str(val) if val is not None else "")
            if row_idx % 2 == 0:
                cell.fill = alt_fill
    
    # Auto column width
    for col in ws.columns:
        max_len = max(len(str(cell.value or "")) for cell in col)
        ws.column_dimensions[col[0].column_letter].width = min(max_len + 4, 40)
    
    # Save to buffer
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