from flask import Flask, render_template , request , redirect, flash, make_response
from config import Database
import os
from werkzeug.utils import secure_filename
from apscheduler.schedulers.background import BackgroundScheduler
import pandas as pd
from datetime import datetime
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
    # Menggunakan file index.html dari folder templates
    return render_template('login.html')

@app.route('/success')
def success():

    # Menggunakan file index.html dari folder templates
    return render_template('success.html')

@app.route('/gambar')
def gambar():
    db = Database
    conn = db.connect_to_database()
    cursor = conn.cursor()
    cursor.execute("SELECT * FROM siswa")
    hasil = cursor.fetchall()
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
    if request.method == 'POST' :
        nama = request.form.get('nama')
        tanggal = request.form.get('tanggal')
        kehadiran = request.form.get('kehadiran')

        db = Database
        conn = db.connect_to_database()
        cursor = conn.cursor()
        cursor.execute('INSERT INTO guru (Name, Date, Status) VALUES (%s, %s, %s) ON DUPLICATE KEY UPDATE Date = VALUES(Date), Status = VALUES(Status)',
        (nama, tanggal, kehadiran))
        conn.commit()
        cursor.close()
        conn.close()
        return redirect('/success')
    # Menggunakan file about.html dari folder templates
    return render_template('guru.html')


@app.route('/murid', methods=['POST', 'GET'])
def about_murid():
    if request.method == 'POST' :
        if 'file' not in request.files:
            flash('No file part')
            return redirect(request.url)

        file = request.files['file']
        if file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        if file:
            filename = secure_filename(file.filename)
            filedata = file.read()
        
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
        # Menggunakan file about.html dari folder templates
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

scheduler = BackgroundScheduler()
scheduler.add_job(export_to_excel, 'cron', hour=20, minute=46)
scheduler.start()

if __name__ == '__main__':
    app.run(debug=True)
