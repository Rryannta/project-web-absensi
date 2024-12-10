from flask import Flask, render_template , request , redirect, flash, make_response
from config import Database # connect file config.py ke app.py / connect database
import os
from werkzeug.utils import secure_filename 
from apscheduler.schedulers.background import BackgroundScheduler
import pandas as pd
from datetime import datetime # library waktu 
import mysql.connector
import pandas as pd



app = Flask(__name__)
app.secret_key = 'raya'
UPLOAD_FOLDER = 'static/uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
EXPORT_FOLDER = 'static/exports'

def connect_to_database():
    conn = mysql.connector.connect(
      host= "localhost",
      user= "root",
      password= "",
      database= "absensi_tkj1"
    )
    return conn

@app.route('/')
def home():
    # Menggunakan file login.html dari folder templates
    return render_template('login.html')

@app.route('/success')
def success():

    # Menggunakan file success.html dari folder templates
    return render_template('success.html')

@app.route('/gambar')
def gambar():
    db = Database
    conn = db.connect_to_database()
    cursor = conn.cursor() # untuk mengeksekusi query sql
    cursor.execute("SELECT * FROM siswa") # memilih semua data dari tabel siswa
    hasil = cursor.fetchall() # menyimpan dan mengambil semua data dari tabel yg sudah dipilih
    conn.close()
    return render_template('image.html', results=hasil) 


@app.route('/image/<int:image_id>') 
def get_image(image_id):
    conn = connect_to_database()
    cursor = conn.cursor()
    cursor.execute("SELECT file_data FROM siswa WHERE id = %s", (image_id,))
    result = cursor.fetchone()
    conn.close()

    if result:
        image_data = result[0]
        response = make_response(image_data)
        response.headers.set('Content-Type', 'image/jpeg')  # Sesuaikan format gambar
        return response
    else:
        return "Image not found", 404
    
@app.route('/guru', methods=['POST', 'GET'])
def about_guru():
    if request.method == 'POST' : # Jika method POST
        nama = request.form.get('nama')
        tanggal = request.form.get('tanggal')
        kehadiran = request.form.get('kehadiran') # diambil dari name yang ada dihtml

        db = Database
        conn = db.connect_to_database()
        cursor = conn.cursor() # untuk mengeksekusi query sql
        cursor.execute('INSERT INTO guru (Name, Date, Status) VALUES (%s, %s, %s) ON DUPLICATE KEY UPDATE Date = VALUES(Date), Status = VALUES(Status)',
        (nama, tanggal, kehadiran)) # memasukkan data ke tabel guru
        conn.commit() # untuk menyimpan perubahan data
        cursor.close()
        conn.close()
        return redirect('/success')
    # Menggunakan file guru.html dari folder templates
    return render_template('guru.html')


@app.route('/murid', methods=['POST', 'GET'])
def about_murid():
    if request.method == 'POST' :
        if 'file' not in request.files: # Jika tidak ada file yang diupload
            flash('No file part')
            return redirect(request.url) # kembali ke halaman sebelumnya

        file = request.files['file'] # mengambil file yang diupload
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        if file:
            filename = secure_filename(file.filename) # mengambil nama file
            filedata = file.read() # membaca isi file
        
        name = request.form.get('name')
        date = request.form.get('date')
        keterangan = request.form.get('keterangan')
        alasan = request.form.get('alasan')
        bukti = request.form.get('image')

        db = Database
        conn = db.connect_to_database()
        cursor = conn.cursor()
        cursor.execute('INSERT INTO siswa (Name, Date, Status, Reason, Proof, file_name, file_data) VALUES (%s, %s, %s, %s, %s, %s, %s) ON DUPLICATE KEY UPDATE Date = VALUES(Date), Status = VALUES(Status),Name = VALUES(Name)', (name, date, keterangan, alasan, bukti, filename, filedata))
        conn.commit()
        cursor.close()
        conn.close()
        return redirect('/success')
        # Menggunakan file murid.html dari folder templates
    return render_template('murid.html')

def export_to_excel():
    """Mengekspor data siswa dan guru ke file Excel dengan nama file sesuai hari."""
    try:
        conn = connect_to_database()

        # Ekspor tabel siswa
        siswa_query = """
        SELECT ID, Name, Date, Status, Reason
        FROM siswa 
        WHERE Date = CURDATE();
        """
        siswa_df = pd.read_sql(siswa_query, conn)

        # Ekspor tabel guru
        guru_query = """
        SELECT ID, Name, Date, Status
        FROM guru 
        WHERE Date = CURDATE();
        """
        guru_df = pd.read_sql(guru_query, conn)

        # Tentukan hari ke berapa
        day_of_month = datetime.now().day
        filename = f"{EXPORT_FOLDER}/day_{day_of_month}.xlsx"

        # Membuat file Excel dengan beberapa sheet
        with pd.ExcelWriter(filename, engine='openpyxl') as writer:
            siswa_df.to_excel(writer, sheet_name='Siswa', index=False)
            guru_df.to_excel(writer, sheet_name='Guru', index=False)

        print(f"Data berhasil diekspor ke {filename}")
        conn.close()
    except Exception as e:
        print(f"Error saat mengekspor data: {e}")

scheduler = BackgroundScheduler() # Membuat scheduler
scheduler.add_job(export_to_excel, 'cron', hour=7, minute=7) # Menjadwalkan ekspor data setiap hari pukul 07:07
scheduler.start() 

if __name__ == '__main__':
    app.run(debug=True)
