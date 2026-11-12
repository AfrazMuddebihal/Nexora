from flask import Flask, render_template, request, redirect, url_for, session, flash, send_file
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3, io, os, json, urllib.request
from datetime import date, datetime
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = os.environ.get("SECRET_KEY", "change-this-secret-on-render")
DB = os.environ.get("DATABASE_PATH", "nexora.db")

def conn():
    c = sqlite3.connect(DB)
    c.row_factory = sqlite3.Row
    c.execute("PRAGMA foreign_keys=ON")
    return c

def init():
    with conn() as c:
        c.executescript("""
        CREATE TABLE IF NOT EXISTS teachers(
          id INTEGER PRIMARY KEY, name TEXT NOT NULL, email TEXT UNIQUE NOT NULL,
          password_hash TEXT NOT NULL, created_at TEXT NOT NULL
        );
        CREATE TABLE IF NOT EXISTS classrooms(
          id INTEGER PRIMARY KEY, teacher_id INTEGER NOT NULL, name TEXT NOT NULL,
          created_at TEXT NOT NULL, FOREIGN KEY(teacher_id) REFERENCES teachers(id)
        );
        CREATE TABLE IF NOT EXISTS students(
          id INTEGER PRIMARY KEY, classroom_id INTEGER NOT NULL, roll TEXT NOT NULL,
          name TEXT NOT NULL, parent_phone TEXT, active INTEGER DEFAULT 1,
          UNIQUE(classroom_id,roll),
          FOREIGN KEY(classroom_id) REFERENCES classrooms(id)
        );
        CREATE TABLE IF NOT EXISTS attendance(
          student_id INTEGER NOT NULL, day TEXT NOT NULL, status TEXT NOT NULL,
          UNIQUE(student_id,day), FOREIGN KEY(student_id) REFERENCES students(id)
        );
        CREATE TABLE IF NOT EXISTS sms_log(
          id INTEGER PRIMARY KEY, student_id INTEGER, day TEXT, status TEXT,
          detail TEXT, created_at TEXT, UNIQUE(student_id,day),
          FOREIGN KEY(student_id) REFERENCES students(id)
        );
        """)
init()

def teacher_id():
    return session.get("teacher_id")

def require_login():
    if not teacher_id():
        flash("Please log in first.")
        return False
    return True

def owned_classroom(cid):
    with conn() as c:
        return c.execute("SELECT * FROM classrooms WHERE id=? AND teacher_id=?", (cid,teacher_id())).fetchone()

@app.route("/")
def home():
    if not teacher_id():
        return redirect(url_for("login"))
    with conn() as c:
        classes=c.execute("SELECT * FROM classrooms WHERE teacher_id=? ORDER BY name",(teacher_id(),)).fetchall()
    return render_template("home.html", classrooms=classes, teacher_name=session.get("teacher_name"))

@app.route("/signup", methods=["GET","POST"])
def signup():
    if request.method=="POST":
        name=request.form.get("name","").strip()
        email=request.form.get("email","").strip().lower()
        password=request.form.get("password","")
        if not name or not email or len(password)<8:
            flash("Enter your name/email and a password of at least 8 characters.")
            return render_template("auth.html", mode="signup")
        try:
            with conn() as c:
                c.execute("INSERT INTO teachers(name,email,password_hash,created_at) VALUES(?,?,?,?)",
                          (name,email,generate_password_hash(password),datetime.now().isoformat()))
            flash("Account created! Please log in.")
            return redirect(url_for("login"))
        except sqlite3.IntegrityError:
            flash("That email already has an account.")
    return render_template("auth.html", mode="signup")

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method=="POST":
        email=request.form.get("email","").strip().lower()
        password=request.form.get("password","")
        with conn() as c:
            t=c.execute("SELECT * FROM teachers WHERE email=?",(email,)).fetchone()
        if t and check_password_hash(t["password_hash"],password):
            session.clear()
            session["teacher_id"]=t["id"]
            session["teacher_name"]=t["name"]
            return redirect(url_for("home"))
        flash("Incorrect email or password.")
    return render_template("auth.html", mode="login")


@app.route("/register")
def register():
    return redirect(url_for("signup"))


@app.route("/students")
def students():
    if not require_login():
        return redirect(url_for("login"))
    return redirect(url_for("home"))


@app.route("/sms-status")
def sms_status():
    if not require_login():
        return redirect(url_for("login"))

    with conn() as c:
        logs = c.execute("""
            SELECT s.name, s.roll, cl.name AS classroom,
                   l.day, l.status, l.detail
            FROM sms_log l
            JOIN students s ON s.id=l.student_id
            JOIN classrooms cl ON cl.id=s.classroom_id
            WHERE cl.teacher_id=?
            ORDER BY l.day DESC, s.roll
        """, (teacher_id(),)).fetchall()

    return render_template("sms_status.html", logs=logs)

@app.route("/logout")
def logout():
    session.clear()
    flash("Logged out.")
    return redirect(url_for("login"))

@app.route("/delete-account", methods=["POST"])
def delete_account():
    if not require_login():
        return redirect(url_for("login"))

    tid = teacher_id()

    with conn() as c:
        c.execute("""
            DELETE FROM sms_log
            WHERE student_id IN (
                SELECT id FROM students
                WHERE classroom_id IN (
                    SELECT id FROM classrooms WHERE teacher_id=?
                )
            )
        """, (tid,))

        c.execute("""
            DELETE FROM attendance
            WHERE student_id IN (
                SELECT id FROM students
                WHERE classroom_id IN (
                    SELECT id FROM classrooms WHERE teacher_id=?
                )
            )
        """, (tid,))

        c.execute("""
            DELETE FROM students
            WHERE classroom_id IN (
                SELECT id FROM classrooms WHERE teacher_id=?
            )
        """, (tid,))

        c.execute("DELETE FROM classrooms WHERE teacher_id=?", (tid,))
        c.execute("DELETE FROM teachers WHERE id=?", (tid,))

    session.clear()
    flash("Your account has been permanently deleted.")
    return redirect(url_for("signup"))

@app.route("/classrooms", methods=["POST"])
def create_classroom():
    if not require_login(): return redirect(url_for("login"))
    name=request.form.get("name","").strip()
    if not name:
        flash("Enter a classroom name.")
    else:
        with conn() as c:
            c.execute("INSERT INTO classrooms(teacher_id,name,created_at) VALUES(?,?,?)",
                      (teacher_id(),name,datetime.now().isoformat()))
        flash("Classroom created.")
    return redirect(url_for("home"))

@app.route("/classroom/<int:cid>/delete", methods=["POST"])
def delete_classroom(cid):
    if not require_login():
        return redirect(url_for("login"))

    tid = teacher_id()

    with conn() as c:
        classroom = c.execute(
            "SELECT id FROM classrooms WHERE id=? AND teacher_id=?",
            (cid, tid)
        ).fetchone()

        if not classroom:
            return "Classroom not found", 404

        c.execute("""
            DELETE FROM sms_log
            WHERE student_id IN (
                SELECT id FROM students WHERE classroom_id=?
            )
        """, (cid,))

        c.execute("""
            DELETE FROM attendance
            WHERE student_id IN (
                SELECT id FROM students WHERE classroom_id=?
            )
        """, (cid,))

        c.execute("DELETE FROM students WHERE classroom_id=?", (cid,))
        c.execute("DELETE FROM classrooms WHERE id=?", (cid,))

    flash("Classroom deleted successfully.")
    return redirect(url_for("home"))

@app.route("/classroom/<int:cid>")
def classroom(cid):
    if not require_login(): return redirect(url_for("login"))
    cl=owned_classroom(cid)
    if not cl: return "Classroom not found",404
    day=request.args.get("day",date.today().isoformat())
    with conn() as c:
        total=c.execute("SELECT count(*) FROM students WHERE classroom_id=? AND active=1",(cid,)).fetchone()[0]
        present=c.execute("SELECT count(*) FROM attendance a JOIN students s ON s.id=a.student_id WHERE s.classroom_id=? AND s.active=1 AND a.day=? AND a.status='Present'",(cid,day)).fetchone()[0]
        absent=c.execute("SELECT count(*) FROM attendance a JOIN students s ON s.id=a.student_id WHERE s.classroom_id=? AND s.active=1 AND a.day=? AND a.status='Absent'",(cid,day)).fetchone()[0]
        rows=c.execute("SELECT s.*,a.status FROM students s LEFT JOIN attendance a ON a.student_id=s.id AND a.day=? WHERE s.classroom_id=? AND s.active=1 ORDER BY s.roll",(day,cid)).fetchall()
        finalized=c.execute("SELECT count(*) FROM sms_log l JOIN students s ON s.id=l.student_id WHERE s.classroom_id=? AND l.day=?",(cid,day)).fetchone()[0]>0
        percentages=c.execute("""SELECT s.id,s.roll,s.name,
          SUM(CASE WHEN a.status='Present' THEN 1 ELSE 0 END) present_count,
          COUNT(a.status) marked_count
          FROM students s LEFT JOIN attendance a ON a.student_id=s.id
          WHERE s.classroom_id=? AND s.active=1 GROUP BY s.id ORDER BY s.roll""",(cid,)).fetchall()
    return render_template("classroom.html",cl=cl,day=day,rows=rows,total=total,present=present,absent=absent,left=max(0,total-present-absent),finalized=finalized,percentages=percentages,overall=(present/total*100 if total else 0))

@app.route("/classroom/<int:cid>/students",methods=["POST"])
def add_student(cid):
    if not require_login(): return redirect(url_for("login"))
    if not owned_classroom(cid): return "Classroom not found",404
    roll=request.form.get("roll","").strip(); name=request.form.get("name","").strip(); phone=request.form.get("parent_phone","").strip()
    if not roll or not name:
        flash("Roll number and student name are required.")
    else:
        try:
            with conn() as c:c.execute("INSERT INTO students(classroom_id,roll,name,parent_phone) VALUES(?,?,?,?)",(cid,roll,name,phone))
            flash("Student added.")
        except sqlite3.IntegrityError: flash("That roll number already exists in this classroom.")
    return redirect(url_for("classroom",cid=cid))

@app.route("/classroom/<int:cid>/mark/<int:sid>/<status>",methods=["POST"])
def mark(cid,sid,status):
    if not require_login(): return redirect(url_for("login"))
    if status not in ("Present","Absent") or not owned_classroom(cid): return "Invalid request",400
    day=request.form.get("day",date.today().isoformat())
    with conn() as c:
        student=c.execute("SELECT id FROM students WHERE id=? AND classroom_id=? AND active=1",(sid,cid)).fetchone()
        finalized=c.execute("SELECT count(*) FROM sms_log WHERE student_id=? AND day=?",(sid,day)).fetchone()[0]
        if not student:return "Student not found",404
        if finalized:flash("This student/date is finalized; reopen before editing.")
        else:c.execute("INSERT INTO attendance(student_id,day,status) VALUES(?,?,?) ON CONFLICT(student_id,day) DO UPDATE SET status=excluded.status",(sid,day,status))
    return redirect(url_for("classroom",cid=cid,day=day))

def send_sms(phone,name,day):
    auth=os.environ.get("MSG91_AUTHKEY"); flow=os.environ.get("MSG91_FLOW_ID"); sender=os.environ.get("MSG91_SENDER")
    if not (auth and flow and sender):return "DEMO","Provider credentials not configured; no SMS sent"
    payload={"flow_id":flow,"sender":sender,"mobiles":phone,"student":name,"date":day}
    req=urllib.request.Request("https://api.msg91.com/api/v5/flow/",data=json.dumps(payload).encode(),headers={"authkey":auth,"Content-Type":"application/json"},method="POST")
    try:
        with urllib.request.urlopen(req,timeout=15) as r:return ("SENT" if r.status in (200,201) else "FAILED"),r.read().decode()[:500]
    except Exception as e:return "FAILED",str(e)[:500]

@app.route("/classroom/<int:cid>/finalize/<day>",methods=["POST"])
def finalize(cid,day):
    if not require_login(): return redirect(url_for("login"))
    if not owned_classroom(cid):return "Classroom not found",404
    with conn() as c:
        total=c.execute("SELECT count(*) FROM students WHERE classroom_id=? AND active=1",(cid,)).fetchone()[0]
        marked=c.execute("SELECT count(*) FROM attendance a JOIN students s ON s.id=a.student_id WHERE s.classroom_id=? AND s.active=1 AND a.day=?",(cid,day)).fetchone()[0]
        if marked<total:flash("Mark every student before finalizing.");return redirect(url_for("classroom",cid=cid,day=day))
        absent=c.execute("SELECT s.id,s.name,s.parent_phone FROM students s JOIN attendance a ON a.student_id=s.id WHERE s.classroom_id=? AND s.active=1 AND a.day=? AND a.status='Absent'",(cid,day)).fetchall()
    for s in absent:
        phone=(s["parent_phone"] or "").strip()
        status,detail=send_sms(phone,s["name"],day) if phone else ("NO NUMBER","Parent mobile number not entered")
        with conn() as c:c.execute("INSERT OR REPLACE INTO sms_log(student_id,day,status,detail,created_at) VALUES(?,?,?,?,?)",(s["id"],day,status,detail,datetime.now().isoformat(timespec="seconds")))
    flash(f"Attendance finalized. Alerts processed for {len(absent)} absent student(s).")
    return redirect(url_for("classroom",cid=cid,day=day))

@app.route("/classroom/<int:cid>/reopen/<day>",methods=["POST"])
def reopen(cid,day):
    if not require_login():return redirect(url_for("login"))
    if not owned_classroom(cid):return "Classroom not found",404
    with conn() as c:c.execute("DELETE FROM sms_log WHERE student_id IN (SELECT id FROM students WHERE classroom_id=?) AND day=?",(cid,day))
    flash("Date reopened.")
    return redirect(url_for("classroom",cid=cid,day=day))

@app.route("/classroom/<int:cid>/export")
def export(cid):
    if not require_login():return redirect(url_for("login"))
    if not owned_classroom(cid):return "Classroom not found",404
    day=request.args.get("day",date.today().isoformat())
    with conn() as c:rows=c.execute("SELECT roll,name,COALESCE(status,'Not marked') FROM students LEFT JOIN attendance ON id=student_id AND day=? WHERE classroom_id=? AND active=1 ORDER BY roll",(day,cid)).fetchall()
    w=Workbook();s=w.active;s.append(["Roll","Name","Status"]);[s.append(list(x)) for x in rows]
    b=io.BytesIO();w.save(b);b.seek(0)
    return send_file(b,as_attachment=True,download_name=f"NEXORA-{day}.xlsx")

if __name__=="__main__":
    app.run(host="0.0.0.0",port=int(os.environ.get("PORT",5000)))
